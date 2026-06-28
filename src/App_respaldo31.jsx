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

/* ============================================================
   SISTEMA DE NOTIFICACIONES (TOASTS) Y CONFIRMACIÓN
   A nivel de módulo para que funcione en cualquier componente.
   Reemplaza los antiguos alert() y window.confirm().
   ============================================================ */

let _toastId = 0;
const _toastSubs = new Set();
const _toastState = { items: [] };

function _emitToasts() {
  _toastSubs.forEach((fn) => fn([..._toastState.items]));
}

// Detecta el tipo de mensaje automáticamente según su contenido.
function _detectarTipoToast(mensaje) {
  const m = String(mensaje || "").toLowerCase();
  if (/no se pudo|no se pudieron|error|falló|fallo|no se encontr|no cuadr|inválid|invalid/.test(m)) return "error";
  if (/correctamente|exitos|guardad|actualizad|eliminad|importad|registrad/.test(m)) return "success";
  if (/debes|ingresa|revisa|falta|advertencia|atención|atencion|ya fue|ya existe/.test(m)) return "warning";
  return "info";
}

// API global. Uso: toast("Mensaje") o toast("Mensaje", "success")
function toast(mensaje, tipo) {
  const id = ++_toastId;
  const t = { id, mensaje: String(mensaje ?? ""), tipo: tipo || _detectarTipoToast(mensaje) };
  _toastState.items = [..._toastState.items, t];
  _emitToasts();
  setTimeout(() => {
    _toastState.items = _toastState.items.filter((x) => x.id !== id);
    _emitToasts();
  }, 4200);
  return id;
}

function ToastContainer() {
  const [items, setItems] = useState([]);
  useEffect(() => {
    _toastSubs.add(setItems);
    setItems([..._toastState.items]);
    return () => { _toastSubs.delete(setItems); };
  }, []);

  const estilos = {
    success: { borde: COLORS.success, icono: "✓" },
    error:   { borde: COLORS.danger,  icono: "✕" },
    warning: { borde: COLORS.warning, icono: "⚠" },
    info:    { borde: COLORS.accent,  icono: "ℹ" },
  };

  return (
    <div style={{ position:"fixed", top:16, right:16, zIndex:99999, display:"flex", flexDirection:"column", gap:10, maxWidth:"min(380px, 92vw)" }}>
      {items.map((t) => {
        const e = estilos[t.tipo] || estilos.info;
        return (
          <div key={t.id} style={{
            background: COLORS.card,
            border: `1px solid ${e.borde}`,
            borderLeft: `4px solid ${e.borde}`,
            borderRadius: 10,
            padding: "12px 14px",
            color: COLORS.text,
            fontSize: 13.5,
            boxShadow: "0 8px 24px rgba(0,0,0,.45)",
            display: "flex",
            alignItems: "flex-start",
            gap: 10,
            animation: "sf-toast-in .18s ease-out",
          }}>
            <span style={{ color: e.borde, fontWeight: 800, fontSize: 15, lineHeight: 1.3 }}>{e.icono}</span>
            <span style={{ flex: 1, lineHeight: 1.35 }}>{t.mensaje}</span>
            <button
              onClick={() => { _toastState.items = _toastState.items.filter((x) => x.id !== t.id); _emitToasts(); }}
              style={{ background:"none", border:"none", color: COLORS.muted, cursor:"pointer", fontSize:14, padding:0 }}
            >✕</button>
          </div>
        );
      })}
      <style>{`@keyframes sf-toast-in { from { opacity:0; transform: translateX(20px); } to { opacity:1; transform: translateX(0); } }`}</style>
    </div>
  );
}

// Confirmación basada en promesa. Uso: const ok = await confirmDialog("¿Seguro?")
let _confirmSub = null;
function confirmDialog(mensaje) {
  return new Promise((resolve) => {
    if (!_confirmSub) {
      // Fallback de seguridad por si el contenedor no está montado.
      resolve(window.confirm(mensaje));
      return;
    }
    _confirmSub({ mensaje: String(mensaje ?? ""), resolve });
  });
}

function ConfirmContainer() {
  const [estado, setEstado] = useState(null);
  useEffect(() => {
    _confirmSub = (data) => setEstado(data);
    return () => { _confirmSub = null; };
  }, []);

  if (!estado) return null;

  const cerrar = (valor) => {
    estado.resolve(valor);
    setEstado(null);
  };

  return (
    <div
      onClick={() => cerrar(false)}
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:100000, padding:16 }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{ background: COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:22, width:"min(420px, 94vw)", color: COLORS.text, boxShadow:"0 14px 40px rgba(0,0,0,.5)" }}
      >
        <div style={{ fontSize:15, lineHeight:1.5, whiteSpace:"pre-line", marginBottom:18 }}>{estado.mensaje}</div>
        <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
          <button
            onClick={() => cerrar(false)}
            style={{ background: COLORS.subtle, border:`1px solid ${COLORS.border}`, color: COLORS.text, borderRadius:8, padding:"9px 16px", fontWeight:700, cursor:"pointer" }}
          >Cancelar</button>
          <button
            onClick={() => cerrar(true)}
            style={{ background: COLORS.danger, border:"none", color:"#fff", borderRadius:8, padding:"9px 16px", fontWeight:700, cursor:"pointer" }}
          >Confirmar</button>
        </div>
      </div>
    </div>
  );
}

// Los datos reales (cotizaciones / notas de venta) se cargan exclusivamente desde Supabase.


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

function estadoProduccionDesdeNota(n) {
  if ((n?.proceso || "").toLowerCase() === "entregado") {
    return { texto: "🚚 Entregada", color: COLORS.success, clave: "entregada" };
  }

  if (n?.media_cana) {
    return { texto: "✅ Terminada", color: COLORS.success, clave: "terminada" };
  }

  const tieneAvance = Boolean(
    n?.mdf_cortado ||
    n?.lamina_cortada ||
    n?.tupizado ||
    n?.armado ||
    n?.pegado ||
    n?.postformado
  );

  if (tieneAvance || String(n?.proceso || "").toLowerCase() === "en producción") {
    return { texto: "🔵 En proceso", color: "#60a5fa", clave: "en_proceso" };
  }

  return { texto: "🟡 Pendiente", color: COLORS.warning, clave: "pendiente" };
}

function calcularProcesoDesdeAvance(n) {
  if ((n?.proceso || "").toLowerCase() === "entregado") return "entregado";
  if (n?.media_cana) return "terminado";
  if (
    n?.mdf_cortado ||
    n?.lamina_cortada ||
    n?.tupizado ||
    n?.armado ||
    n?.pegado ||
    n?.postformado
  ) return "en producción";
  return "en espera";
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

// ============= MODAL DE LOGIN ADMIN =============
function LoginAdminModal({ open, onClose, usuario, setUsuario, password, setPassword, onLogin, cargando }) {
  if (!open) return null;

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,.7)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 100000,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.border}`,
          borderRadius: 14,
          padding: 28,
          width: "min(380px, 92vw)",
          color: COLORS.text,
        }}
      >
        <h2 style={{ margin: "0 0 20px", color: COLORS.accent, fontSize: 18 }}>
          🔐 Acceso Administrador
        </h2>

        <label style={{ display: "block", marginBottom: 14, fontSize: 13, color: COLORS.muted }}>
          Usuario
        </label>
        <input
          type="text"
          value={usuario}
          onChange={(e) => setUsuario(e.target.value)}
          placeholder="Admin"
          disabled={cargando}
          style={{
            width: "100%",
            padding: "10px 12px",
            borderRadius: 8,
            border: `1px solid ${COLORS.border}`,
            background: COLORS.surface,
            color: COLORS.text,
            boxSizing: "border-box",
            marginBottom: 14,
            fontSize: 13,
          }}
        />

        <label style={{ display: "block", marginBottom: 14, fontSize: 13, color: COLORS.muted }}>
          Contraseña
        </label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          placeholder="••••••••"
          disabled={cargando}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !cargando) onLogin();
          }}
          style={{
            width: "100%",
            padding: "10px 12px",
            borderRadius: 8,
            border: `1px solid ${COLORS.border}`,
            background: COLORS.surface,
            color: COLORS.text,
            boxSizing: "border-box",
            marginBottom: 18,
            fontSize: 13,
          }}
        />

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button
            onClick={onClose}
            disabled={cargando}
            style={{
              background: COLORS.subtle,
              border: `1px solid ${COLORS.border}`,
              color: COLORS.text,
              borderRadius: 8,
              padding: "9px 16px",
              fontWeight: 700,
              cursor: cargando ? "not-allowed" : "pointer",
              opacity: cargando ? 0.6 : 1,
            }}
          >
            Cancelar
          </button>
          <button
            onClick={onLogin}
            disabled={cargando}
            style={{
              background: COLORS.accent,
              border: "none",
              color: "#0f0e0c",
              borderRadius: 8,
              padding: "9px 16px",
              fontWeight: 700,
              cursor: cargando ? "not-allowed" : "pointer",
              opacity: cargando ? 0.6 : 1,
            }}
          >
            {cargando ? "Verificando..." : "Ingresar"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ============= MODAL PARA REGISTRAR NUEVO ERROR =============
function NuevoErrorModal({
  open,
  onClose,
  tiposError,
  formError,
  setFormError,
  busquedaNV,
  setBusquedaNV,
  nvsDisponibles,
  onGuardar,
}) {
  if (!open) return null;

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,.65)",
        display: "flex",
        alignItems: "flex-start",
        justifyContent: "center",
        zIndex: 9999,
        overflowY: "auto",
        padding: "20px 10px",
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.border}`,
          borderRadius: 14,
          padding: 22,
          width: "min(460px, 92vw)",
          maxHeight: "90vh",
          overflowY: "auto",
          color: COLORS.text,
          boxSizing: "border-box",
        }}
      >
        <h2 style={{ marginTop: 0, color: COLORS.accent }}>📋 Registrar Error de Producción</h2>

        <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
          Nota de Venta
        </label>
        <div style={{ position: "relative", marginBottom: 14 }}>
          <input
            type="text"
            placeholder="Busca por número NV..."
            value={busquedaNV}
            onChange={(e) => setBusquedaNV(e.target.value)}
            style={{
              width: "100%",
              padding: "10px 12px",
              borderRadius: 8,
              border: `1px solid ${COLORS.border}`,
              background: COLORS.surface,
              color: COLORS.text,
              boxSizing: "border-box",
              fontSize: 13,
            }}
          />
          {busquedaNV && nvsDisponibles.length > 0 && (
            <div
              style={{
                position: "absolute",
                top: "100%",
                left: 0,
                right: 0,
                background: COLORS.surface,
                border: `1px solid ${COLORS.border}`,
                borderTop: "none",
                borderRadius: "0 0 8px 8px",
                maxHeight: 150,
                overflowY: "auto",
                zIndex: 10000,
              }}
            >
              {nvsDisponibles.map((nv) => (
                <div
                  key={nv.id}
                  onClick={() => {
                    setFormError((prev) => ({ ...prev, nv: String(nv.numero) }));
                    setBusquedaNV("");
                  }}
                  style={{
                    padding: "10px 12px",
                    cursor: "pointer",
                    borderBottom: `1px solid ${COLORS.border}`,
                    fontSize: 13,
                  }}
                  onMouseEnter={(e) => (e.target.style.background = COLORS.card)}
                  onMouseLeave={(e) => (e.target.style.background = "transparent")}
                >
                  <strong>{nv.numero}</strong> — {nv.cliente}
                </div>
              ))}
            </div>
          )}
        </div>
        {formError.nv && <div style={{ fontSize: 11, color: COLORS.success, marginBottom: 10 }}>✓ {formError.nv}</div>}

        <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
          Tipo de Error
        </label>
        <select
          value={formError.tipo}
          onChange={(e) => setFormError((prev) => ({ ...prev, tipo: e.target.value }))}
          style={{
            width: "100%",
            padding: "10px 12px",
            borderRadius: 8,
            border: `1px solid ${COLORS.border}`,
            background: COLORS.surface,
            color: COLORS.text,
            boxSizing: "border-box",
            marginBottom: 14,
            fontSize: 13,
          }}
        >
          <option value="">Selecciona un tipo...</option>
          {tiposError.map((tipo) => (
            <option key={tipo} value={tipo}>
              {tipo}
            </option>
          ))}
        </select>

        <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
          Descripción (opcional)
        </label>
        <textarea
          value={formError.descripcion}
          onChange={(e) => setFormError((prev) => ({ ...prev, descripcion: e.target.value }))}
          placeholder="Ej: Se rompió la lámina durante el transporte, llegó con grieta de fábrica..."
          style={{
            width: "100%",
            minHeight: 70,
            padding: "10px 12px",
            borderRadius: 8,
            border: `1px solid ${COLORS.border}`,
            background: COLORS.surface,
            color: COLORS.text,
            boxSizing: "border-box",
            marginBottom: 14,
            fontSize: 13,
            fontFamily: "inherit",
            resize: "none",
          }}
        />

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
          <div>
            <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
              Días de Retrabaljo
            </label>
            <input
              type="number"
              min="0"
              value={formError.dias}
              onChange={(e) => setFormError((prev) => ({ ...prev, dias: Number(e.target.value) }))}
              style={{
                width: "100%",
                padding: "10px 12px",
                borderRadius: 8,
                border: `1px solid ${COLORS.border}`,
                background: COLORS.surface,
                color: COLORS.text,
                boxSizing: "border-box",
                fontSize: 13,
              }}
            />
          </div>

          <div>
            <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
              Monto Material Perdido
            </label>
            <input
              type="number"
              min="0"
              value={formError.monto}
              onChange={(e) => setFormError((prev) => ({ ...prev, monto: Number(e.target.value) }))}
              placeholder="0"
              style={{
                width: "100%",
                padding: "10px 12px",
                borderRadius: 8,
                border: `1px solid ${COLORS.border}`,
                background: COLORS.surface,
                color: COLORS.text,
                boxSizing: "border-box",
                fontSize: 13,
              }}
            />
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button
            onClick={onClose}
            style={{
              background: COLORS.subtle,
              border: `1px solid ${COLORS.border}`,
              color: COLORS.text,
              borderRadius: 8,
              padding: "9px 16px",
              fontWeight: 700,
              cursor: "pointer",
              fontSize: 13,
            }}
          >
            Cancelar
          </button>
          <button
            onClick={onGuardar}
            style={{
              background: COLORS.success,
              border: "none",
              color: "#fff",
              borderRadius: 8,
              padding: "9px 16px",
              fontWeight: 700,
              cursor: "pointer",
              fontSize: 13,
            }}
          >
            Guardar Error
          </button>
        </div>
      </div>
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
    media_cana: nota.media_cana || false,
  });

  const ordenProcesos = [
    "mdf_cortado",
    "lamina_cortada",
    "tupizado",
    "armado",
    "pegado",
    "postformado",
    "media_cana",
  ];

  const nombresProcesos = {
    mdf_cortado: "MDF cortado",
    lamina_cortada: "Lámina cortada",
    tupizado: "Tupizado",
    armado: "Armado",
    pegado: "Pegado",
    postformado: "Postformado",
    media_cana: "Media caña",
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
function GestionNVModal({ nota, abonos = [], detalles = [], onClose, onSave }) {
  const [nuevoAbono, setNuevoAbono] = useState("");
  const [observacionAbono, setObservacionAbono] = useState("");
  const [fechaAbono, setFechaAbono] = useState(new Date().toISOString().split("T")[0]);
  const [medioPagoAbono, setMedioPagoAbono] = useState("transferencia");
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
    fechaAbono,
    medioPagoAbono,
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
  const estadoProd = estadoProduccionDesdeNota(nota);

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
        <div style={{
          margin:"10px 0 14px",
          padding:10,
          borderRadius:10,
          border:`1px solid ${estadoProd.color}`,
          background:estadoProd.color + "18",
          color:estadoProd.color,
          fontWeight:800
        }}>
          Estado producción: {estadoProd.texto}
        </div>

        <div style={{
          background:COLORS.surface,
          border:`1px solid ${COLORS.border}`,
          borderRadius:10,
          padding:12,
          margin:"10px 0 14px"
        }}>
          <h3 style={{ margin:"0 0 10px", color:COLORS.accent, fontSize:15 }}>
            Detalle del pedido
          </h3>

          {detalles.length === 0 ? (
            <p style={{ margin:0, color:COLORS.muted, fontSize:13 }}>
              Esta nota de venta no tiene detalle importado.
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

<div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
  <label>Fecha abono
    <input
      type="date"
      value={fechaAbono}
      onChange={(e) => setFechaAbono(e.target.value)}
      style={{
        width:"100%",
        padding:10,
        margin:"6px 0 12px",
        borderRadius:8,
        border:`1px solid ${COLORS.border}`,
        background:COLORS.surface,
        color:COLORS.text,
        boxSizing:"border-box"
      }}
    />
  </label>

  <label>Forma de pago
    <select
      value={medioPagoAbono}
      onChange={(e) => setMedioPagoAbono(e.target.value)}
      style={{
        width:"100%",
        padding:10,
        margin:"6px 0 12px",
        borderRadius:8,
        border:`1px solid ${COLORS.border}`,
        background:COLORS.surface,
        color:COLORS.text,
        boxSizing:"border-box"
      }}
    >
      <option value="transferencia">Transferencia</option>
      <option value="efectivo">Efectivo</option>
      <option value="cheque">Cheque</option>
    </select>
  </label>
</div>

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
            {a.medio_pago ? `${String(a.medio_pago).charAt(0).toUpperCase()}${String(a.medio_pago).slice(1)}` : "Sin medio"}
            {a.observacion ? ` · ${a.observacion}` : ""}
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

function EditarNVCompletaModal({ nota, detalles = [], onClose, onSave }) {
  const [form, setForm] = useState({
    numero: nota.numero || "",
    cotizacion: nota.cotizacion || "",
    cliente: nota.cliente || "",
    fecha: nota.fecha || "",
    total: nota.total || ""
  });

  const normalizarDetalle = (d, idx) => ({
    material: d.material || "",
    cantidad: d.cantidad ?? "",
    descripcion: d.descripcion || d.tipo || "",
    alto: d.alto ?? d.largo ?? "",
    ancho: d.ancho ?? "",
    color: d.color || "",
    orden: d.orden || idx + 1
  });

  const [filas, setFilas] = useState(
    detalles.length > 0
      ? detalles.map(normalizarDetalle)
      : [{ material:"", cantidad:"", descripcion:"", alto:"", ancho:"", color:"", orden:1 }]
  );

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", cerrarConEsc);
    return () => window.removeEventListener("keydown", cerrarConEsc);
  }, [onClose]);

  const cambiarForm = (campo, valor) => {
    setForm(prev => ({ ...prev, [campo]: valor }));
  };

  const cambiarFila = (idx, campo, valor) => {
    setFilas(prev => prev.map((f, i) => i === idx ? { ...f, [campo]: valor } : f));
  };

  const agregarFila = () => {
    setFilas(prev => [...prev, { material:"", cantidad:"", descripcion:"", alto:"", ancho:"", color:"", orden:prev.length + 1 }]);
  };

  const eliminarFila = (idx) => {
    setFilas(prev => prev.length === 1 ? prev : prev.filter((_, i) => i !== idx).map((f, i) => ({ ...f, orden:i + 1 })));
  };

  const guardar = () => {
    if (!String(form.numero || "").trim()) {
      toast("Debes ingresar el número de la Nota de Venta.");
      return;
    }

    if (!String(form.cliente || "").trim()) {
      toast("Debes ingresar el cliente.");
      return;
    }

    const detallesLimpios = filas
      .map((f, idx) => ({
        material: String(f.material || "").trim(),
        cantidad: Number(f.cantidad || 0),
        descripcion: String(f.descripcion || "").trim(),
        alto: Number(f.alto || 0),
        ancho: Number(f.ancho || 0),
        color: String(f.color || "").trim(),
        orden: idx + 1
      }))
      .filter(f => f.material || f.descripcion || f.cantidad || f.alto || f.ancho || f.color);

    onSave({
      notaOriginal: nota,
      notaEditada: {
        ...nota,
        numero: String(form.numero).trim(),
        cotizacion: String(form.cotizacion || "").trim(),
        cliente: String(form.cliente).trim(),
        fecha: form.fecha,
        total: Number(form.total || 0)
      },
      detallesEditados: detallesLimpios
    });
  };

  const inputStyle = {
    width:"100%",
    boxSizing:"border-box",
    padding:"9px 10px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    fontSize:13
  };

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"18px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:18,
          width:"min(1050px, 96vw)",
          maxHeight:"90vh",
          overflow:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, marginBottom:14 }}>
          <div>
            <h2 style={{ margin:0, color:COLORS.accent }}>Editar Nota de Venta</h2>
            <div style={{ fontSize:12, color:COLORS.muted }}>Corrige datos leídos desde Excel sin perder abonos ni gestión.</div>
          </div>
          <button onClick={onClose} style={{ background:COLORS.danger, color:"#fff", border:"none", borderRadius:8, padding:"8px 11px", cursor:"pointer", fontWeight:700 }}>Cerrar</button>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(170px, 1fr))", gap:10, marginBottom:14 }}>
          <label style={{ fontSize:12, color:COLORS.muted }}>Número NV<input value={form.numero} onChange={e=>cambiarForm("numero", e.target.value)} style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>N° Cotización asociada<input value={form.cotizacion} onChange={e=>cambiarForm("cotizacion", e.target.value)} placeholder="Ej: 12143" style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>Cliente<input value={form.cliente} onChange={e=>cambiarForm("cliente", e.target.value)} style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>Fecha<input type="date" value={form.fecha || ""} onChange={e=>cambiarForm("fecha", e.target.value)} style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>Total<input type="number" value={form.total} onChange={e=>cambiarForm("total", e.target.value)} style={inputStyle}/></label>
        </div>

        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", margin:"8px 0 10px", gap:10 }}>
          <h3 style={{ margin:0, color:COLORS.success, fontSize:15 }}>Detalle del pedido</h3>
          <button onClick={agregarFila} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"8px 12px", cursor:"pointer", fontWeight:700 }}>+ Agregar fila</button>
        </div>

        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:850, fontSize:12 }}>
            <thead>
              <tr style={{ background:COLORS.surface }}>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Cantidad</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Material</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Detalle / Tipo</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Largo / Alto</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Ancho</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Color</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"center" }}>Eliminar</th>
              </tr>
            </thead>
            <tbody>
              {filas.map((fila, idx) => (
                <tr key={idx}>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.cantidad} onChange={e=>cambiarFila(idx,"cantidad",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input value={fila.material} onChange={e=>cambiarFila(idx,"material",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input value={fila.descripcion} onChange={e=>cambiarFila(idx,"descripcion",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.alto} onChange={e=>cambiarFila(idx,"alto",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.ancho} onChange={e=>cambiarFila(idx,"ancho",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input value={fila.color} onChange={e=>cambiarFila(idx,"color",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}`, textAlign:"center" }}>
                    <button onClick={()=>eliminarFila(idx)} disabled={filas.length===1} style={{ background:filas.length===1?COLORS.subtle:COLORS.danger, color:"#fff", border:"none", borderRadius:7, padding:"7px 9px", cursor:filas.length===1?"not-allowed":"pointer" }}>🗑</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={{ display:"flex", justifyContent:"flex-end", gap:10, marginTop:16 }}>
          <button onClick={onClose} style={{ padding:"10px 14px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, cursor:"pointer" }}>Cancelar</button>
          <button onClick={guardar} style={{ padding:"10px 14px", borderRadius:8, border:"none", background:COLORS.success, color:"#fff", fontWeight:700, cursor:"pointer" }}>Guardar cambios</button>
        </div>
      </div>
    </div>
  );
}


function EditarCotizacionCompletaModal({ cotizacion, detalles = [], onClose, onSave }) {
  const [form, setForm] = useState({
    numero: cotizacion.numero || "",
    cliente: cotizacion.cliente || "",
    fecha: cotizacion.fecha || "",
    total: cotizacion.total || ""
  });

  const normalizarDetalle = (d, idx) => ({
    unidad: d.unidad ?? d.cantidad ?? "",
    tipo: d.tipo || d.descripcion || "",
    largo: d.largo ?? d.alto ?? "",
    ancho: d.ancho ?? "",
    color: d.color || "",
    valor: d.valor ?? "",
    total: d.total ?? "",
    orden: d.orden || idx + 1
  });

  const [filas, setFilas] = useState(
    detalles.length > 0
      ? detalles.map(normalizarDetalle)
      : [{ unidad:"", tipo:"", largo:"", ancho:"", color:"", valor:"", total:"", orden:1 }]
  );

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", cerrarConEsc);
    return () => window.removeEventListener("keydown", cerrarConEsc);
  }, [onClose]);

  const inputStyle = {
    width:"100%",
    boxSizing:"border-box",
    padding:"9px 10px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    fontSize:13
  };

  const cambiarForm = (campo, valor) => setForm(prev => ({ ...prev, [campo]: valor }));
  const cambiarFila = (idx, campo, valor) => setFilas(prev => prev.map((f, i) => i === idx ? { ...f, [campo]: valor } : f));
  const agregarFila = () => setFilas(prev => [...prev, { unidad:"", tipo:"", largo:"", ancho:"", color:"", valor:"", total:"", orden:prev.length + 1 }]);
  const eliminarFila = (idx) => setFilas(prev => prev.length === 1 ? prev : prev.filter((_, i) => i !== idx).map((f, i) => ({ ...f, orden:i + 1 })));

  const guardar = () => {
    if (!String(form.numero || "").trim()) {
      toast("Debes ingresar el número de la cotización.");
      return;
    }
    if (!String(form.cliente || "").trim()) {
      toast("Debes ingresar el cliente.");
      return;
    }

    const detallesLimpios = filas
      .map((f, idx) => ({
        unidad: Number(f.unidad || 0),
        tipo: String(f.tipo || "").trim(),
        largo: Number(f.largo || 0),
        ancho: Number(f.ancho || 0),
        color: String(f.color || "").trim(),
        valor: Number(f.valor || 0),
        total: Number(f.total || 0),
        orden: idx + 1
      }))
      .filter(f => f.unidad || f.tipo || f.largo || f.ancho || f.color || f.valor || f.total);

    onSave({
      cotizacionOriginal: cotizacion,
      cotizacionEditada: {
        ...cotizacion,
        numero: String(form.numero).trim(),
        cliente: String(form.cliente).trim(),
        fecha: form.fecha,
        total: Number(form.total || 0)
      },
      detallesEditados: detallesLimpios
    });
  };

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"18px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:18,
          width:"min(1050px, 96vw)",
          maxHeight:"90vh",
          overflow:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, marginBottom:14 }}>
          <div>
            <h2 style={{ margin:0, color:COLORS.accent }}>Editar Cotización</h2>
            <div style={{ fontSize:12, color:COLORS.muted }}>Corrige datos leídos desde Excel.</div>
          </div>
          <button onClick={onClose} style={{ background:COLORS.danger, color:"#fff", border:"none", borderRadius:8, padding:"8px 11px", cursor:"pointer", fontWeight:700 }}>Cerrar</button>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(170px, 1fr))", gap:10, marginBottom:14 }}>
          <label style={{ fontSize:12, color:COLORS.muted }}>Número cotización<input value={form.numero} onChange={e=>cambiarForm("numero", e.target.value)} style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>Cliente<input value={form.cliente} onChange={e=>cambiarForm("cliente", e.target.value)} style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>Fecha<input type="date" value={form.fecha || ""} onChange={e=>cambiarForm("fecha", e.target.value)} style={inputStyle}/></label>
          <label style={{ fontSize:12, color:COLORS.muted }}>Total<input type="number" value={form.total} onChange={e=>cambiarForm("total", e.target.value)} style={inputStyle}/></label>
        </div>

        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", margin:"8px 0 10px", gap:10 }}>
          <h3 style={{ margin:0, color:COLORS.success, fontSize:15 }}>Detalle de cotización</h3>
          <button onClick={agregarFila} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"8px 12px", cursor:"pointer", fontWeight:700 }}>+ Agregar fila</button>
        </div>

        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:900, fontSize:12 }}>
            <thead>
              <tr style={{ background:COLORS.surface }}>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Cantidad</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Tipo</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Largo</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Ancho</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Color</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Valor</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"left" }}>Total</th>
                <th style={{ padding:7, border:`1px solid ${COLORS.border}`, textAlign:"center" }}>Eliminar</th>
              </tr>
            </thead>
            <tbody>
              {filas.map((fila, idx) => (
                <tr key={idx}>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.unidad} onChange={e=>cambiarFila(idx,"unidad",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input value={fila.tipo} onChange={e=>cambiarFila(idx,"tipo",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.largo} onChange={e=>cambiarFila(idx,"largo",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.ancho} onChange={e=>cambiarFila(idx,"ancho",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input value={fila.color} onChange={e=>cambiarFila(idx,"color",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.valor} onChange={e=>cambiarFila(idx,"valor",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}` }}><input type="number" value={fila.total} onChange={e=>cambiarFila(idx,"total",e.target.value)} style={inputStyle}/></td>
                  <td style={{ padding:5, border:`1px solid ${COLORS.border}`, textAlign:"center" }}>
                    <button onClick={()=>eliminarFila(idx)} disabled={filas.length===1} style={{ background:filas.length===1?COLORS.subtle:COLORS.danger, color:"#fff", border:"none", borderRadius:7, padding:"7px 9px", cursor:filas.length===1?"not-allowed":"pointer" }}>🗑</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={{ display:"flex", justifyContent:"flex-end", gap:10, marginTop:16 }}>
          <button onClick={onClose} style={{ padding:"10px 14px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, cursor:"pointer" }}>Cancelar</button>
          <button onClick={guardar} style={{ padding:"10px 14px", borderRadius:8, border:"none", background:COLORS.success, color:"#fff", fontWeight:700, cursor:"pointer" }}>Guardar cambios</button>
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
function PreviewOCModal({ productos, documentos = [], onClose, onConfirm }) {
  const hoy = new Date().toISOString().split("T")[0];
  const [fecha, setFecha] = useState(hoy);
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

  const documentoPrincipal = documentos.length === 1 ? documentos[0]?.documento : "OC MASIVA";

  const confirmar = () => {
    onConfirm({
      fecha,
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
          width:"680px",
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
          La app detectó automáticamente el número de OC desde la celda F13 de cada archivo. Si alguna OC ya fue importada, se omite para no duplicar stock.
        </p>

        <div style={{
          background:COLORS.surface,
          border:`1px solid ${COLORS.border}`,
          borderRadius:10,
          padding:12,
          margin:"12px 0"
        }}>
          <div style={{ display:"flex", justifyContent:"space-between", gap:10, alignItems:"center", marginBottom:8 }}>
            <b style={{ color:COLORS.text }}>Documentos detectados</b>
            <span style={{ color:COLORS.accent, fontWeight:800 }}>{documentos.length}</span>
          </div>

          {documentos.length === 0 ? (
            <div style={{ color:COLORS.muted, fontSize:13 }}>No hay documentos detectados.</div>
          ) : (
            <div style={{ display:"grid", gap:6, maxHeight:130, overflowY:"auto" }}>
              {documentos.map((doc, idx) => (
                <div key={idx} style={{ display:"grid", gridTemplateColumns:"90px 1fr", gap:8, fontSize:13 }}>
                  <b style={{ color:COLORS.accent }}>{doc.documento}</b>
                  <span style={{ color:COLORS.muted, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{doc.archivo}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(160px, 1fr))", gap:10, margin:"12px 0" }}>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Fecha contable</label>
            <input type="date" value={fecha} onChange={(e)=>setFecha(e.target.value)} style={inputStyle} />
          </div>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Proveedor opcional</label>
            <input value={proveedor} onChange={(e)=>setProveedor(e.target.value)} placeholder="Ej: Imperial" style={inputStyle} />
          </div>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Total compra IVA incluido</label>
            <input type="number" value={totalCompra} onChange={(e)=>setTotalCompra(e.target.value)} placeholder={documentos.length > 1 ? "Opcional" : "Ej: 119000"} style={inputStyle} />
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
              border:`1px solid ${estadoPago === "pendiente" ? COLORS.warning : COLORS.border}`,
              background:estadoPago === "pendiente" ? COLORS.warning : COLORS.surface,
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

function FacturaXmlModal({ factura, onClose, onConfirm }) {
  const [modo, setModo] = useState("inventario_contabilidad");
  const [estadoPago, setEstadoPago] = useState("pagado");
  const [cuentaCompra, setCuentaCompra] = useState("Compras");

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);
    return () => window.removeEventListener("keydown", cerrarConEsc);
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
    onConfirm({ modo, estadoPago, cuentaCompra });
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
          width:"720px",
          maxWidth:"95%",
          maxHeight:"90vh",
          overflowY:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, marginBottom:12 }}>
          <div>
            <h2 style={{ margin:"0 0 4px", color:COLORS.accent }}>Vista previa Factura XML</h2>
            <div style={{ fontSize:12, color:COLORS.muted }}>
              {factura.documento} · {factura.proveedor}
            </div>
          </div>
          <button onClick={onClose} style={{ background:COLORS.danger, color:"#fff", border:"none", borderRadius:8, padding:"8px 11px", cursor:"pointer", fontWeight:700 }}>Cerrar</button>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(160px, 1fr))", gap:10, marginBottom:14 }}>
          <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:10 }}>
            <div style={{ fontSize:11, color:COLORS.muted }}>Fecha</div>
            <b>{factura.fecha}</b>
          </div>
          <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:10 }}>
            <div style={{ fontSize:11, color:COLORS.muted }}>RUT proveedor</div>
            <b>{factura.rutProveedor}</b>
          </div>
          <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:10 }}>
            <div style={{ fontSize:11, color:COLORS.muted }}>Neto</div>
            <b>{fmt(factura.neto)}</b>
          </div>
          <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:10 }}>
            <div style={{ fontSize:11, color:COLORS.muted }}>IVA CF</div>
            <b>{fmt(factura.iva)}</b>
          </div>
          <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:10 }}>
            <div style={{ fontSize:11, color:COLORS.muted }}>Total</div>
            <b style={{ color:COLORS.accent }}>{fmt(factura.total)}</b>
          </div>
        </div>

        <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:12, marginBottom:14 }}>
          <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent, marginBottom:8 }}>¿Qué hacer con esta factura?</div>
          <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
            <button
              onClick={() => { setModo("inventario_contabilidad"); setCuentaCompra("Compras"); }}
              style={{ padding:"9px 12px", borderRadius:8, border:`1px solid ${modo === "inventario_contabilidad" ? COLORS.success : COLORS.border}`, background:modo === "inventario_contabilidad" ? COLORS.success : COLORS.card, color:modo === "inventario_contabilidad" ? "#fff" : COLORS.text, fontWeight:700, cursor:"pointer" }}
            >
              Inventario + Contabilidad
            </button>
            <button
              onClick={() => { setModo("solo_contabilidad"); setCuentaCompra("Gastos generales"); }}
              style={{ padding:"9px 12px", borderRadius:8, border:`1px solid ${modo === "solo_contabilidad" ? COLORS.warning : COLORS.border}`, background:modo === "solo_contabilidad" ? COLORS.warning : COLORS.card, color:modo === "solo_contabilidad" ? "#111" : COLORS.text, fontWeight:700, cursor:"pointer" }}
            >
              Solo Contabilidad
            </button>
          </div>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(180px, 1fr))", gap:10, marginBottom:14 }}>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Estado de pago</label>
            <select value={estadoPago} onChange={(e)=>setEstadoPago(e.target.value)} style={inputStyle}>
              <option value="pagado">Pagada con banco</option>
              <option value="caja">Pagada con caja</option>
              <option value="pendiente">Pendiente proveedor</option>
            </select>
          </div>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Cuenta contable</label>
            <select value={cuentaCompra} onChange={(e)=>setCuentaCompra(e.target.value)} style={inputStyle}>
              <option value="Compras">Compras / Materiales</option>
              <option value="Materiales de oficina">Materiales de oficina</option>
              <option value="Combustible">Combustible</option>
              <option value="Servicios">Servicios</option>
              <option value="Fletes">Fletes</option>
              <option value="Arriendo">Arriendo</option>
              <option value="Gastos generales">Gastos generales</option>
            </select>
          </div>
        </div>

        <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent, marginBottom:8 }}>Detalle del XML</div>
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {factura.detalles.map((p, index) => (
            <div key={index} style={{ border:`1px solid ${COLORS.border}`, borderRadius:8, padding:10, background:COLORS.surface, display:"grid", gridTemplateColumns:"1fr auto auto", gap:10, alignItems:"center" }}>
              <span>{p.nombre}</span>
              <b style={{ color:COLORS.accent }}>{p.cantidad}</b>
              <span style={{ color:COLORS.muted }}>{fmt(p.total)}</span>
            </div>
          ))}
        </div>

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end", marginTop:16 }}>
          <button onClick={onClose} style={{ padding:"9px 12px", borderRadius:8 }}>Cancelar</button>
          <button onClick={confirmar} style={{ padding:"9px 14px", borderRadius:8, border:"none", background:COLORS.success, color:"#fff", fontWeight:700, cursor:"pointer" }}>
            Confirmar importación
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
      toast("Debes ingresar el nombre del producto.");
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
      toast("Debes ingresar el nombre del producto.");
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
        toast("Debes ingresar una fecha.");
        return;
      }
      if (!form.detalle.trim()) {
        toast("Debes ingresar el detalle/cuenta del asiento.");
        return;
      }

      const debe = Number(form.debe) || 0;
      const haber = Number(form.haber) || 0;

      if (debe <= 0 && haber <= 0) {
        toast("Debes ingresar un monto en Debe o Haber.");
        return;
      }

      if (debe > 0 && haber > 0) {
        toast("Un asiento no debe tener Debe y Haber al mismo tiempo. Deja uno de los dos en cero.");
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
      toast("Debes ingresar al menos 2 líneas para un asiento manual.");
      return;
    }

    for (const fila of filasValidas) {
      if (!fila.fecha) {
        toast("Todas las filas con datos deben tener fecha.");
        return;
      }
      if (!fila.detalle) {
        toast("Todas las filas con datos deben tener cuenta/detalle.");
        return;
      }
      if (fila.debe <= 0 && fila.haber <= 0) {
        toast("Cada fila debe tener monto en Debe o en Haber.");
        return;
      }
      if (fila.debe > 0 && fila.haber > 0) {
        toast("Una fila no puede tener Debe y Haber al mismo tiempo.");
        return;
      }
    }

    if (!estaCuadrado) {
      toast("El asiento no está cuadrado. El total Debe debe ser igual al total Haber.");
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
      toast("Debes ingresar una fecha.");
      return null;
    }
    if (total <= 0) {
      toast("Debes ingresar un monto mayor a cero.");
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
        toast("El origen y destino no pueden ser iguales.");
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

  const normalizar = (txt) => String(txt || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();

  const esNumeroUtil = (v) => {
    const n = Number(v);
    return Number.isFinite(n) && n > 0;
  };

  const obtenerIndex = (encabezados, palabras) => {
    for (let i = 0; i < encabezados.length; i++) {
      const h = normalizar(encabezados[i]);
      if (palabras.some(p => h === p || h.includes(p))) return i;
    }
    return -1;
  };

  let inicioDetalle = -1;
  let columnas = null;

  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i];
    const textoFila = fila.map(normalizar).join(" ");

    if (
      textoFila.includes("MATERIAL") &&
      textoFila.includes("CANTIDAD") &&
      (textoFila.includes("DESCRIPCION") || textoFila.includes("DESCRIPCIÓN")) &&
      textoFila.includes("COLOR")
    ) {
      const materialIdx = obtenerIndex(fila, ["MATERIAL"]);
      const cantidadIdx = obtenerIndex(fila, ["CANTIDAD", "UNIDAD"]);
      const descripcionIdx = obtenerIndex(fila, ["DESCRIPCION", "DESCRIPCIÓN", "TIPO"]);
      const altoIdx = obtenerIndex(fila, ["ALTO", "LARGO"]);
      const anchoIdx = obtenerIndex(fila, ["ANCHO"]);
      const colorIdx = obtenerIndex(fila, ["COLOR"]);

      if (cantidadIdx >= 0 && descripcionIdx >= 0 && altoIdx >= 0 && anchoIdx >= 0) {
        inicioDetalle = i + 1;
        columnas = { materialIdx, cantidadIdx, descripcionIdx, altoIdx, anchoIdx, colorIdx };
        break;
      }
    }
  }

  if (inicioDetalle === -1 || !columnas) return [];

  const detalles = [];

  for (let i = inicioDetalle; i < filas.length; i++) {
    const fila = filas[i];

    const material = columnas.materialIdx >= 0 ? String(fila[columnas.materialIdx] || "").trim() : "";
    const cantidad = fila[columnas.cantidadIdx];
    const descripcion = String(fila[columnas.descripcionIdx] || "").trim();
    const alto = fila[columnas.altoIdx];
    const ancho = fila[columnas.anchoIdx];
    const color = columnas.colorIdx >= 0 ? String(fila[columnas.colorIdx] || "").trim() : "";

    const filaVacia =
      !material &&
      !cantidad &&
      !descripcion &&
      !alto &&
      !ancho &&
      !color;

    if (filaVacia) break;

    const descripcionNorm = normalizar(descripcion);

    if (
      descripcionNorm.includes("DESPACHO") ||
      descripcionNorm.includes("FLETE") ||
      descripcionNorm.includes("TRASLADO")
    ) {
      continue;
    }

    if (!esNumeroUtil(cantidad) && !descripcion) continue;
    if (!esNumeroUtil(alto) && !esNumeroUtil(ancho)) continue;

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
  const obtenerFechaEntregaDesdeResumen = (sheet) => {
    // En las NV/Barranes actuales, la fecha límite de entrega viene en G11.
    const valorG11 = sheet?.["G11"]?.v;
    return valorG11 ? excelDateToISO(valorG11) : "";
  };

  const procesarCotizacionArchivo = async (file) => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets["RESUMEN"];

  if (!sheet) {
    throw new Error(`No existe hoja RESUMEN en ${file.name}`);
  }

  const nombresHojas = workbook.SheetNames;

  const hojaDetalleNombre = nombresHojas.find(
    n =>
      n !== "RESUMEN" &&
      n !== "NOTA DE VENTA" &&
      n !== "PRODUCCION" &&
      n !== "SEGUIMIENTO"
  );

  const detalles = [];

  if (hojaDetalleNombre && workbook.Sheets[hojaDetalleNombre]) {
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
  }

  const numero = sheet["A2"]?.v;
  const cliente = sheet["B2"]?.v;
  const fecha = sheet["C2"]?.v;
  const total = sheet["D2"]?.v;

  if (!numero || !total) {
    throw new Error(`No se pudo leer número o total en ${file.name}`);
  }

  const ok = await importarCotizacionExcel({
    numero,
    cliente,
    fecha,
    total,
    detalles
  });

  return ok;
};

  const importarCotizacion = async (event) => {
  const files = Array.from(event.target.files || []);

  if (files.length === 0) return;

  let importadas = 0;
  let saltadas = 0;
  const errores = [];

  for (const file of files) {
    try {
      const ok = await procesarCotizacionArchivo(file);
      if (ok) importadas += 1;
      else saltadas += 1;
    } catch (error) {
      console.error(error);
      errores.push(`${file.name}: ${error.message || "Error desconocido"}`);
    }
  }

  event.target.value = "";

  toast(
    `Importación de cotizaciones terminada.\n\nImportadas: ${importadas}\nDuplicadas o saltadas: ${saltadas}\nCon error: ${errores.length}` +
    (errores.length ? `\n\nErrores:\n${errores.slice(0, 5).join("\n")}` : "")
  );

  window.location.reload();
};
  const procesarNotaVentaArchivo = async (file) => {
    const data = await file.arrayBuffer();

    const workbook = XLSX.read(data);

    const sheet = workbook.Sheets['RESUMEN'];

    if (!sheet) {
      throw new Error(`No existe hoja RESUMEN en ${file.name}`);
    }

    const cotizacion = sheet['A2']?.v;
    const notaVenta = sheet['B2']?.v;
    const cliente = sheet['C2']?.v;
    const fecha = sheet['D2']?.v;
    const total = sheet['E2']?.v;
    const fechaEntrega = obtenerFechaEntregaDesdeResumen(sheet);

    if (!notaVenta || !total) {
      throw new Error(`No se pudo leer número o total de NV en ${file.name}`);
    }

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
        .from("notas_venta")
        .update({
          tipo_documento: "nv",
          fecha_entrega_estimada: fechaEntrega || null
        })
        .eq("numero", notaVenta);

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
          throw new Error(`La NV ${notaVenta} se importó, pero falló el detalle de producción.`);
        }
      }

      await crearAsientoVentaNVAutomatica({
        fecha: excelDateToISO(fecha),
        numero: notaVenta,
        cliente,
        totalVenta: total
      });
    }

    return ok;
  };

  const importarExcel = async (event) => {
    const files = Array.from(event.target.files || []);

    if (files.length === 0) return;

    let importadas = 0;
    let saltadas = 0;
    const errores = [];

    for (const file of files) {
      try {
        const ok = await procesarNotaVentaArchivo(file);
        if (ok) importadas += 1;
        else saltadas += 1;
      } catch (error) {
        console.error(error);
        errores.push(`${file.name}: ${error.message || "Error desconocido"}`);
      }
    }

    event.target.value = "";

    toast(
      `Importación de notas de venta terminada.\n\nImportadas: ${importadas}\nDuplicadas o saltadas: ${saltadas}\nCon error: ${errores.length}` +
      (errores.length ? `\n\nErrores:\n${errores.slice(0, 5).join("\n")}` : "")
    );

    window.location.reload();
  };

  const procesarBarranArchivo = async (file) => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets["RESUMEN"];

    if (!sheet) {
      throw new Error(`No existe hoja RESUMEN en ${file.name}`);
    }

    const cotizacion = sheet["A2"]?.v;
    const barran = sheet["B2"]?.v;
    const cliente = sheet["C2"]?.v;
    const fecha = sheet["D2"]?.v;
    const total = sheet["E2"]?.v;
    const fechaEntrega = obtenerFechaEntregaDesdeResumen(sheet);

    if (!barran || !total) {
      throw new Error(`No se pudo leer número o total de Barrán en ${file.name}`);
    }

    const { data: existente, error: errorBuscar } = await supabase
      .from("notas_venta")
      .select("id")
      .eq("numero", barran)
      .eq("tipo_documento", "barran")
      .maybeSingle();

    if (errorBuscar) {
      console.error(errorBuscar);
      throw new Error(`No se pudo verificar si el Barrán ${barran} ya existe.`);
    }

    if (existente?.id) {
      return false;
    }

    const { data: nuevoBarran, error: errorInsert } = await supabase
      .from("notas_venta")
      .insert({
        numero: Number(barran) || barran,
        cliente: String(cliente || ""),
        fecha: excelDateToISO(fecha),
        total: Number(total || 0),
        cotizacion_id: null,
        tipo_documento: "barran",
        estado_pago: "pendiente",
        materiales: "falta",
        proceso: "en espera",
        fecha_entrega_estimada: fechaEntrega || null
      })
      .select("*")
      .single();

    if (errorInsert) {
      console.error(errorInsert);
      throw new Error(`No se pudo importar el Barrán ${barran}.`);
    }

    const detallesProduccion = obtenerDetalleProduccionDesdeExcel(workbook);

    await supabase
      .from("detalles_notas_venta_produccion")
      .delete()
      .eq("nota_venta_numero", String(barran));

    if (detallesProduccion.length > 0) {
      const detallesParaGuardar = detallesProduccion.map((d) => ({
        nota_venta_numero: String(barran),
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
        throw new Error(`El Barrán ${barran} se importó, pero falló el detalle de producción.`);
      }
    }

    setNotas(prev => [{
      id: `supabase-${nuevoBarran.id}`,
      numero: String(nuevoBarran.numero),
      cliente: nuevoBarran.cliente || "(sin cliente)",
      fecha: nuevoBarran.fecha?.split("T")[0] || new Date().toISOString().split("T")[0],
      total: nuevoBarran.total || 0,
      cotizacion: null,
      abono: nuevoBarran.abono || 0,
      estado_pago: nuevoBarran.estado_pago || "pendiente",
      materiales: nuevoBarran.materiales || "falta",
      proceso: nuevoBarran.proceso || "en espera",
      observaciones: nuevoBarran.observaciones || "",
      fecha_entrega_estimada: nuevoBarran.fecha_entrega_estimada || "",
      produccion_observaciones: nuevoBarran.produccion_observaciones || "",
      mdf_cortado: nuevoBarran.mdf_cortado || false,
      lamina_cortada: nuevoBarran.lamina_cortada || false,
      tupizado: nuevoBarran.tupizado || false,
      armado: nuevoBarran.armado || false,
      pegado: nuevoBarran.pegado || false,
      postformado: nuevoBarran.postformado || false,
      media_cana: nuevoBarran.media_cana || false,
      tipo_documento: "barran"
    }, ...prev]);

    return true;
  };

  const importarBarranesExcel = async (event) => {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;

    let importadas = 0;
    let saltadas = 0;
    const errores = [];

    for (const file of files) {
      try {
        const ok = await procesarBarranArchivo(file);
        if (ok) importadas += 1;
        else saltadas += 1;
      } catch (error) {
        console.error(error);
        errores.push(`${file.name}: ${error.message || "Error desconocido"}`);
      }
    }

    event.target.value = "";

    toast(
      `Importación de barranes terminada.\n\nImportados: ${importadas}\nDuplicados o saltados: ${saltadas}\nCon error: ${errores.length}` +
      (errores.length ? `\n\nErrores:\n${errores.slice(0, 5).join("\n")}` : "")
    );

    window.location.reload();
  };
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

    const { data: dataNotasDirectas, error: errorNotasDirectas } = await supabase
      .from("notas_venta")
      .select("*, cotizaciones(numero, cliente, total)")
      .order("fecha", { ascending: false });

    const dataNotas = errorNotasDirectas ? await obtenerNotasVenta() : (dataNotasDirectas || []);

    const notasFormateadas = dataNotas.map((n) => ({
      id: `supabase-${n.id}`,
      numero: String(n.numero),
      cliente: n.cliente || n.cotizaciones?.cliente || "(sin cliente)",
      fecha: n.fecha?.split("T")[0] || new Date().toISOString().split("T")[0],
     total: n.total || n.cotizaciones?.total || 0,
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
      media_cana: n.media_cana || false,
      tipo_documento: n.tipo_documento || "nv",
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
  const detallesNVData = dataDetallesNV || [];
  setDetallesNotasVenta(detallesNVData);
  setNotas(prev => prev.map(n => {
    if (n.cotizacion) return n;
    const detalleRelacionado = detallesNVData.find(d => String(d.nota_venta_numero) === String(n.numero));
    return detalleRelacionado?.cotizacion_numero
      ? { ...n, cotizacion: String(detalleRelacionado.cotizacion_numero) }
      : n;
  }));
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
const { data: dataDocumentosImportados, error: errorDocumentosImportados } = await supabase
  .from("documentos_importados")
  .select("*")
  .order("created_at", { ascending: false });

if (!errorDocumentosImportados) {
  setDocumentosImportados(dataDocumentosImportados || []);
}

const { data: dataCartolaBancaria, error: errorCartolaBancaria } = await supabase
  .from("cartola_bancaria")
  .select("*")
  .order("fecha", { ascending: false })
  .order("created_at", { ascending: false });

if (!errorCartolaBancaria) {
  setCartolaMovimientos(dataCartolaBancaria || []);
}
const { data: dataCuentasPagar, error: errorCuentasPagar } = await supabase
  .from("cuentas_pagar")
  .select("*")
  .order("fecha_vencimiento", { ascending: true });

if (!errorCuentasPagar) {
  setCuentasPagar(dataCuentasPagar || []);
}

const { data: dataAbonosCuentas, error: errorAbonosCuentas } = await supabase
  .from("cuentas_pagar_abonos")
  .select("*")
  .order("fecha", { ascending: false });

if (!errorAbonosCuentas) {
  setAbonosCuentasPagar(dataAbonosCuentas || []);
}

const { data: dataTrabajadores, error: errorTrabajadores } = await supabase
  .from("trabajadores")
  .select("*")
  .order("nombre", { ascending: true });

if (!errorTrabajadores) {
  setTrabajadores(dataTrabajadores || []);
}

const { data: dataDocumentosTrabajadores, error: errorDocumentosTrabajadores } = await supabase
  .from("trabajador_documentos")
  .select("*")
  .order("created_at", { ascending: false });

if (!errorDocumentosTrabajadores) {
  setDocumentosTrabajadores(dataDocumentosTrabajadores || []);
}
  }

  cargarDatos()
    .catch((e) => { console.error(e); toast("Hubo un error al cargar los datos.", "error"); })
    .finally(() => setCargando(false));
}, []);
  const [tab, setTab] = useState("produccion");
  const [esAdmin, setEsAdmin] = useState(false);
  const [usuarioActual, setUsuarioActual] = useState(null); // {username, nombre, secciones}
  const [modalLogin, setModalLogin] = useState(false);
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [cargandoLogin, setCargandoLogin] = useState(false);
  const [erroresProduccion, setErroresProduccion] = useState([]);
  const [modalNuevoError, setModalNuevoError] = useState(false);
  const [formError, setFormError] = useState({ nv: "", tipo: "", descripcion: "", dias: 0, monto: 0 });
  const [busquedaNV, setBusquedaNV] = useState("");
  const [errorEditando, setErrorEditando] = useState(null);
  const [formEditError, setFormEditError] = useState({ dias: 0, monto: 0 });
  const [clienteSeleccionado, setClienteSeleccionado] = useState(null);
  const [busquedaClienteAnalisis, setBusquedaClienteAnalisis] = useState("");
  const [filtroFechasAnalisis, setFiltroFechasAnalisis] = useState("mes_actual");
  const [filtroResumen, setFiltroResumen] = useState("mes_actual");
  const [filter, setFilter] = useState("");
  const [fechaDesde, setFechaDesde] = useState("");
  const [fechaHasta, setFechaHasta] = useState("");
  const [mesFiltro, setMesFiltro] = useState(new Date().toISOString().slice(0, 7));
  const [showVencidas, setShowVencidas] = useState(false);
  const [seguimiento, setSeguimiento] = useState({});
  const [cargando, setCargando] = useState(true);
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
  const [previewOCDocumentos, setPreviewOCDocumentos] = useState([]);
  const [modalPreviewOC, setModalPreviewOC] = useState(false);
  const [documentosImportados, setDocumentosImportados] = useState([]);
  const [previewFacturaXml, setPreviewFacturaXml] = useState(null);
  const [modalFacturaXml, setModalFacturaXml] = useState(false);
  const [modalSobrantesLaminado, setModalSobrantesLaminado] = useState(null);
  const [busquedaLaminadoNombre, setBusquedaLaminadoNombre] = useState("");
  const [busquedaSobranteLargo, setBusquedaSobranteLargo] = useState("");
  const [busquedaSobranteAncho, setBusquedaSobranteAncho] = useState("");
  const [notas, setNotas] = useState([]);
  const [abonosNV, setAbonosNV] = useState([]);
  const [modalCot, setModalCot] = useState(null);
  const [modalEditarCotizacion, setModalEditarCotizacion] = useState(null);
  const [modalNV, setModalNV] = useState(null);
  const [modalEditarNV, setModalEditarNV] = useState(null);
  const [modalProduccion, setModalProduccion] = useState(null);
  const [filtroProduccion, setFiltroProduccion] = useState("todos");
  const [ordenProduccion, setOrdenProduccion] = useState("correlativo");
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
  const [mesCxc, setMesCxc] = useState(new Date().toISOString().slice(0, 7));
  const [busquedaContabilidad, setBusquedaContabilidad] = useState("");
  const [cartolaMovimientos, setCartolaMovimientos] = useState([]);
  const [mesCartola, setMesCartola] = useState(new Date().toISOString().slice(0, 7));
  const mesActual = new Date().toISOString().slice(0, 7);

  const cambiarMesSimple = (mes, delta) => {
  const [year, month] = mes.split("-").map(Number);
  const fecha = new Date(year, month - 1 + delta, 1);
  return `${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, "0")}`;
};
  const [busquedaCartola, setBusquedaCartola] = useState("");
  const [modalClasificarCartola, setModalClasificarCartola] = useState(null);
  const [categoriaCartolaSeleccionada, setCategoriaCartolaSeleccionada] = useState("");
  const [busquedaCategoriaCartola, setBusquedaCategoriaCartola] = useState("");
  const [busquedaCxcCartola, setBusquedaCxcCartola] = useState("");
  const [cxcSeleccionadaCartola, setCxcSeleccionadaCartola] = useState(null);
  const [tipoAbonoCartola, setTipoAbonoCartola] = useState("abono");
  const [clientesCreditoAutorizado, setClientesCreditoAutorizado] = useState([]);
  const [mesCalendario, setMesCalendario] = useState(new Date().toISOString().slice(0, 7));
  const [cuentasPagar, setCuentasPagar] = useState([]);
  const [abonosCuentasPagar, setAbonosCuentasPagar] = useState([]);
  const [modalCuentaPagar, setModalCuentaPagar] = useState(false);
  const [modalAbonoCuentaPagar, setModalAbonoCuentaPagar] = useState(null);
  const [modalAbonoCxc, setModalAbonoCxc] = useState(null);
  const [trabajadores, setTrabajadores] = useState([]);
  const [documentosTrabajadores, setDocumentosTrabajadores] = useState([]);
  const [modalTrabajador, setModalTrabajador] = useState(false);
  const [trabajadorArchivo, setTrabajadorArchivo] = useState(null);
  

  useEffect(() => {
    try {
      const guardados = localStorage.getItem("sf-clientes-credito-autorizado");
      if (guardados) setClientesCreditoAutorizado(JSON.parse(guardados));
    } catch (e) {
      console.warn("No se pudo leer clientes con crédito autorizado", e);
    }
  }, []);

  // Cargar seguimiento desde Supabase (con fallback a localStorage).
  // Tabla esperada: seguimiento_cotizaciones (columnas: numero text, entries jsonb)
  useEffect(() => {
    let activo = true;
    (async () => {
      // 1) Cache local inmediato para que la UI no quede vacía mientras llega Supabase.
      try {
        const saved = localStorage.getItem("sf-seguimiento");
        if (saved && activo) setSeguimiento(JSON.parse(saved));
      } catch (e) {}

      // 2) Fuente de verdad: Supabase.
      try {
        const { data, error } = await supabase
          .from("seguimiento_cotizaciones")
          .select("numero, entries");

        if (!error && data && activo) {
          const mapa = {};
          data.forEach((row) => {
            mapa[String(row.numero)] = Array.isArray(row.entries) ? row.entries : [];
          });
          setSeguimiento(mapa);
          try { localStorage.setItem("sf-seguimiento", JSON.stringify(mapa)); } catch (e) {}
        }
      } catch (e) {
        // Si la tabla no existe todavía, se sigue usando el cache local sin romper la app.
        console.warn("seguimiento_cotizaciones no disponible, usando cache local.", e);
      }
    })();
    return () => { activo = false; };
  }, []);

  const saveSeguimiento = useCallback(async (numero, entries) => {
    const updated = { ...seguimiento, [numero]: entries };
    setSeguimiento(updated);
    // Cache local (respaldo / offline).
    try { localStorage.setItem("sf-seguimiento", JSON.stringify(updated)); } catch (e) {}
    // Persistir en Supabase.
    try {
      const { error } = await supabase
        .from("seguimiento_cotizaciones")
        .upsert({ numero: String(numero), entries }, { onConflict: "numero" });
      if (error) {
        console.error(error);
        toast("El seguimiento se guardó localmente, pero no se pudo sincronizar con la nube.", "warning");
      }
    } catch (e) {
      console.error(e);
    }
  }, [seguimiento]);

  // ============= AUTENTICACIÓN DE USUARIOS =============
  // Lista de secciones que requieren rol admin (control total)
  const SECCIONES_ADMIN = ["control_calidad", "cartola", "cxc", "contabilidad", "rrhh"];

  // Verificar sesión guardada al cargar
  useEffect(() => {
    const sesion = sessionStorage.getItem("usuario_session");
    if (sesion) {
      try {
        const u = JSON.parse(sesion);
        if (u && u.username) {
          setUsuarioActual(u);
          // esAdmin si tiene acceso a "all" o a alguna sección admin
          const tieneAll = Array.isArray(u.secciones) && u.secciones.includes("all");
          setEsAdmin(tieneAll);
        }
      } catch (e) {
        console.error("Sesión corrupta:", e);
        sessionStorage.removeItem("usuario_session");
      }
    }
  }, []);

  // Guarda de seguridad: si el tab actual no está permitido, redirigir
  useEffect(() => {
    if (!usuarioActual) return;
    const secs = usuarioActual.secciones || [];
    if (secs.includes("all")) return; // acceso total
    if (!secs.includes(tab)) {
      setTab(secs[0] || "produccion");
    }
  }, [usuarioActual, tab]);

  // Cargar errores de producción si es admin
  useEffect(() => {
    if (!esAdmin) return;
    (async () => {
      try {
        const { data, error } = await supabase
          .from("errores_produccion")
          .select("*")
          .order("fecha_reporte", { ascending: false });
        if (!error) {
          setErroresProduccion(data || []);
        }
      } catch (e) {
        console.error("Error cargando errores de producción:", e);
      }
    })();
  }, [esAdmin]);

  // Login de usuario
  const hacerLoginAdmin = async () => {
    if (!loginUser.trim() || !loginPass.trim()) {
      toast("Ingresa usuario y contraseña", "warning");
      return;
    }
    setCargandoLogin(true);
    try {
      const { data, error } = await supabase
        .from("usuarios")
        .select("username, password, nombre, secciones, activo")
        .eq("username", loginUser.trim())
        .maybeSingle();

      if (error || !data) {
        toast("Usuario o contraseña incorrectos", "error");
        setCargandoLogin(false);
        return;
      }

      if (data.activo === false) {
        toast("Este usuario está desactivado", "error");
        setCargandoLogin(false);
        return;
      }

      if (data.password !== loginPass) {
        toast("Usuario o contraseña incorrectos", "error");
        setCargandoLogin(false);
        return;
      }

      // Éxito
      const usuario = {
        username: data.username,
        nombre: data.nombre || data.username,
        secciones: Array.isArray(data.secciones) ? data.secciones : [],
      };
      sessionStorage.setItem("usuario_session", JSON.stringify(usuario));
      setUsuarioActual(usuario);
      const tieneAll = usuario.secciones.includes("all");
      setEsAdmin(tieneAll);
      setModalLogin(false);
      setLoginUser("");
      setLoginPass("");

      // Llevar al usuario a su primera sección permitida
      const primeraSeccion = usuario.secciones.includes("all")
        ? "produccion"
        : (usuario.secciones[0] || "produccion");
      setTab(primeraSeccion);

      toast(`Bienvenido, ${usuario.nombre}`, "success");
    } catch (e) {
      console.error(e);
      toast("Error al verificar credenciales", "error");
    } finally {
      setCargandoLogin(false);
    }
  };

  // Logout
  const hacerLogoutAdmin = () => {
    sessionStorage.removeItem("usuario_session");
    setUsuarioActual(null);
    setEsAdmin(false);
    setLoginUser("");
    setLoginPass("");
    toast("Sesión cerrada", "info");
  };

  // ============= GESTIÓN DE ERRORES DE PRODUCCIÓN =============
  const tiposError = [
    "Lámina rota (manipulación)",
    "Lámina defectuosa",
    "Lámina quebrada al postformado",
    "Lámina quemada en postformado",
    "Error en el tupizado",
    "Error en la media caña",
    "Error en el enchapado",
    "Otro",
  ];

  // Agregar error
  const guardarNuevoError = async () => {
    if (!formError.nv.trim()) {
      toast("Selecciona una Nota de Venta", "warning");
      return;
    }
    if (!formError.tipo) {
      toast("Selecciona un tipo de error", "warning");
      return;
    }

    const nvSeleccionada = notas.find((n) => n.numero === formError.nv);
    if (!nvSeleccionada) {
      toast("NV no encontrada", "error");
      return;
    }

    const idReal = Number(String(nvSeleccionada.id).replace("supabase-", ""));

    try {
      const { data, error } = await supabase
        .from("errores_produccion")
        .insert([
          {
            nota_venta_id: idReal,
            nota_venta_numero: String(formError.nv),
            cliente: nvSeleccionada.cliente || "",
            tipo_error: formError.tipo,
            descripcion: formError.descripcion || "",
            dias_retrabaljo: Number(formError.dias) || 0,
            monto_material_perdido: Number(formError.monto) || 0,
            reportado_por: "Admin",
          },
        ])
        .select();

      if (error) {
        console.error(error);
        toast("Error al guardar el error de producción", "error");
        return;
      }

      setErroresProduccion((prev) => [data[0], ...prev]);
      setModalNuevoError(false);
      setFormError({ nv: "", tipo: "", descripcion: "", dias: 0, monto: 0 });
      setBusquedaNV("");
      toast("Error de producción registrado", "success");
    } catch (e) {
      console.error(e);
      toast("Error inesperado", "error");
    }
  };

  // Editar error
  const editarError = async (errorId, updates) => {
    try {
      const { data, error } = await supabase
        .from("errores_produccion")
        .update(updates)
        .eq("id", errorId)
        .select();

      if (error) {
        console.error(error);
        toast("Error al actualizar", "error");
        return;
      }

      setErroresProduccion((prev) =>
        prev.map((e) => (e.id === errorId ? data[0] : e))
      );
      toast("Error actualizado", "success");
    } catch (e) {
      console.error(e);
      toast("Error inesperado", "error");
    }
  };

  // Eliminar error
  const eliminarError = async (errorId) => {
    const confirmar = await confirmDialog(
      "¿Seguro que quieres eliminar este registro de error?"
    );
    if (!confirmar) return;

    try {
      const { error } = await supabase
        .from("errores_produccion")
        .delete()
        .eq("id", errorId);

      if (error) {
        console.error(error);
        toast("Error al eliminar", "error");
        return;
      }

      setErroresProduccion((prev) => prev.filter((e) => e.id !== errorId));
      toast("Error eliminado", "success");
    } catch (e) {
      console.error(e);
      toast("Error inesperado", "error");
    }
  };

  // ============= ANÁLISIS DE CLIENTE =============
  // Obtener rango de fechas según filtro
  const obtenerRangoFechas = () => {
    const hoy = new Date();
    let desde = new Date();
    
    switch(filtroFechasAnalisis) {
      case "mes_actual":
        desde.setDate(1);
        break;
      case "ultimos_3_meses":
        desde.setMonth(hoy.getMonth() - 3);
        break;
      case "ultimo_año":
        desde.setFullYear(hoy.getFullYear() - 1);
        break;
      case "año_actual":
        desde = new Date(hoy.getFullYear(), 0, 1);
        break;
      default:
        desde.setDate(1);
    }
    return { desde, hasta: hoy };
  };

  // Filtrar cotizaciones y notas por cliente y fecha
  const analizarCliente = (nombreCliente) => {
    if (!nombreCliente || !cotizaciones || !notasVenta) return null;
    
    const { desde, hasta } = obtenerRangoFechas();
    
    const cotizacionesCliente = (cotizaciones || []).filter((c) => {
      if (!c || !c.cliente || !c.fecha) return false;
      const fecha = new Date(c.fecha);
      return c.cliente.toLowerCase() === nombreCliente.toLowerCase() && 
             fecha >= desde && fecha <= hasta;
    });

    const notasCliente = (notasVenta || []).filter((n) => {
      if (!n || !n.cliente || !n.fecha) return false;
      const fecha = new Date(n.fecha);
      return n.cliente.toLowerCase() === nombreCliente.toLowerCase() && 
             fecha >= desde && fecha <= hasta;
    });

    // Totales
    const totalCotizado = cotizacionesCliente.reduce((s, c) => s + (Number(c.total) || 0), 0);
    const totalVendido = notasCliente.reduce((s, n) => s + (Number(n.total) || 0), 0);
    const tasaConversion = cotizacionesCliente.length > 0 
      ? ((notasCliente.length / cotizacionesCliente.length) * 100).toFixed(1)
      : 0;

    // Frecuencia de compra
    let frecuenciaPromedio = "-";
    if (notasCliente.length > 1) {
      const fechas = notasCliente.map(n => new Date(n.fecha)).sort((a,b) => a - b);
      const diferencias = [];
      for (let i = 1; i < fechas.length; i++) {
        diferencias.push((fechas[i] - fechas[i-1]) / (1000 * 60 * 60 * 24));
      }
      const promedio = (diferencias.reduce((a,b) => a+b, 0) / diferencias.length).toFixed(0);
      frecuenciaPromedio = `Cada ${promedio} días`;
    }

    // Análisis de productos (de detalles de notas de venta)
    const detallesCliente = (detallesNotasVenta || []).filter((d) => {
      if (!d || !d.nota_venta_numero) return false;
      const nv = notasCliente.find(n => String(n.numero) === String(d.nota_venta_numero));
      return nv !== undefined;
    });

    // Tipos de cubierta más comprados
    const tiposCubierta = {};
    detallesCliente.forEach((d) => {
      const tipo = String(d.descripcion || "Sin especificar").trim().toUpperCase();
      tiposCubierta[tipo] = (tiposCubierta[tipo] || 0) + (Number(d.cantidad) || 1);
    });
    const topCubiertas = Object.entries(tiposCubierta)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5)
      .map(([tipo, cant]) => ({ tipo, cantidad: cant }));

    // Medidas más pedidas
    const medidas = {};
    detallesCliente.forEach((d) => {
      if (d.alto && d.ancho) {
        const medida = `${d.alto}x${d.ancho}`;
        medidas[medida] = (medidas[medida] || 0) + (Number(d.cantidad) || 1);
      }
    });
    const topMedidas = Object.entries(medidas)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5)
      .map(([medida, cant]) => ({ medida, cantidad: cant }));

    // Colores más solicitados
    const colores = {};
    detallesCliente.forEach((d) => {
      if (d.color) {
        const color = String(d.color).trim().toUpperCase();
        colores[color] = (colores[color] || 0) + (Number(d.cantidad) || 1);
      }
    });
    const topColores = Object.entries(colores)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5)
      .map(([color, cant]) => ({ color, cantidad: cant }));

    return {
      totalCotizado,
      nroCotizaciones: cotizacionesCliente.length,
      totalVendido,
      nroNotas: notasCliente.length,
      tasaConversion,
      frecuenciaPromedio,
      ultimaCompra: notasCliente.length > 0 
        ? { fecha: notasCliente[0].fecha, monto: notasCliente[0].total }
        : null,
      topCubiertas,
      topMedidas,
      topColores,
      cotizacionesCliente,
      notasCliente,
    };
  };

  // Filtrar NVs para autocomplete
  const nvsDisponibles = notas.filter((n) =>
    String(n.numero).toLowerCase().includes(busquedaNV.toLowerCase())
  );

  const notasVenta = notas.filter(n => (n.tipo_documento || "nv") !== "barran");
  const barranes = notas.filter(n => (n.tipo_documento || "nv") === "barran");

  // Clientes únicos para búsqueda (análisis de cliente)
  const clientesUnicos = cotizaciones && notasVenta 
    ? [...new Set([...cotizaciones, ...notasVenta].map(x => x.cliente).filter(c => c))].sort()
    : [];
  const clientesFiltrados = clientesUnicos.filter(c =>
    c && c.toLowerCase().includes(busquedaClienteAnalisis.toLowerCase())
  );
  const analisisClienteSeleccionado = clienteSeleccionado && cotizaciones && notasVenta && detallesNotasVenta
    ? analizarCliente(clienteSeleccionado)
    : null;

  const convertedNums = new Set(notasVenta.filter(s => s.cotizacion).map(s => s.cotizacion));
  const sinCotizacion = notasVenta.filter(s => !s.cotizacion);
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
  const totalSold = notasVenta.reduce((s,n) => s + n.total, 0);
  const totalActiva = [...activas,...urgentes].reduce((s,q) => s + q.total, 0);
  const totalVencida = vencidas.reduce((s,q) => s + q.total, 0);
  const totalQuoted = cotizaciones.reduce((s,q) => s + q.total, 0);

  const obtenerFechaValida = (fecha) => {
    if (!fecha) return null;
    const f = new Date(`${fecha}T00:00:00`);
    if (!Number.isNaN(f.getTime())) return f;
    const fallback = new Date(fecha);
    return Number.isNaN(fallback.getTime()) ? null : fallback;
  };

  const nombreMesResumen = (mes) => {
    if (!mes) return "Todos los meses";
    const [year, month] = String(mes).split("-");
    const nombre = new Date(Number(year), Number(month) - 1, 1)
      .toLocaleDateString("es-CL", { month:"long", year:"numeric" });
    return nombre.charAt(0).toUpperCase() + nombre.slice(1);
  };

  const inicioMes = (fecha) => new Date(fecha.getFullYear(), fecha.getMonth(), 1);
  const finMes = (fecha) => new Date(fecha.getFullYear(), fecha.getMonth() + 1, 0, 23, 59, 59, 999);
  const sumarMeses = (fecha, meses) => new Date(fecha.getFullYear(), fecha.getMonth() + meses, 1);

  const obtenerRangoResumen = (tipo = filtroResumen, offset = 0) => {
    const hoy = new Date();
    const base = sumarMeses(hoy, offset);

    if (tipo === "todo") {
      return { desde:null, hasta:null, label:"Todo el histórico" };
    }

    if (tipo === "mes_actual") {
      return { desde:inicioMes(base), hasta:finMes(base), label:nombreMesResumen(`${base.getFullYear()}-${String(base.getMonth()+1).padStart(2,"0")}`) };
    }

    if (tipo === "mes_anterior") {
      const mes = sumarMeses(base, -1);
      return { desde:inicioMes(mes), hasta:finMes(mes), label:nombreMesResumen(`${mes.getFullYear()}-${String(mes.getMonth()+1).padStart(2,"0")}`) };
    }

    if (tipo === "ultimos_3_meses") {
      const hasta = finMes(base);
      const desde = inicioMes(sumarMeses(base, -2));
      return { desde, hasta, label:"Últimos 3 meses" };
    }

    if (tipo === "ultimos_6_meses") {
      const hasta = finMes(base);
      const desde = inicioMes(sumarMeses(base, -5));
      return { desde, hasta, label:"Últimos 6 meses" };
    }

    if (tipo === "año_actual") {
      return { desde:new Date(hoy.getFullYear(), 0, 1), hasta:new Date(hoy.getFullYear(), 11, 31, 23, 59, 59, 999), label:`Año ${hoy.getFullYear()}` };
    }

    return { desde:null, hasta:null, label:"Todo el histórico" };
  };

  const estaEnRangoResumen = (fecha, rango) => {
    if (!rango?.desde && !rango?.hasta) return true;
    const f = obtenerFechaValida(fecha);
    if (!f) return false;
    if (rango.desde && f < rango.desde) return false;
    if (rango.hasta && f > rango.hasta) return false;
    return true;
  };

  const calcularMetricasResumen = (rango) => {
    const cotizacionesPeriodo = cotizaciones.filter(c => estaEnRangoResumen(c.fecha, rango));
    const notasPeriodo = notasVenta.filter(n => estaEnRangoResumen(n.fecha, rango));
    const barranesPeriodo = barranes.filter(n => estaEnRangoResumen(n.fecha, rango));
    const ventasPeriodo = [...notasPeriodo, ...barranesPeriodo];
    const convertedPeriodo = new Set(ventasPeriodo.filter(n => n.cotizacion).map(n => n.cotizacion));
    const cotizacionesConEstado = cotizacionesPeriodo.map(q => ({ ...q, status: getStatus(q, convertedPeriodo) }));
    const vendidasPeriodo = cotizacionesConEstado.filter(q => q.status === "vendida");
    const activasPeriodo = cotizacionesConEstado.filter(q => q.status === "activa");
    const urgentesPeriodo = cotizacionesConEstado.filter(q => q.status === "urgente");
    const vencidasPeriodo = cotizacionesConEstado.filter(q => q.status === "vencida");
    const montoCotizado = cotizacionesPeriodo.reduce((s,q) => s + Number(q.total || 0), 0);
    const montoVendido = ventasPeriodo.reduce((s,n) => s + Number(n.total || 0), 0);
    const conversion = cotizacionesPeriodo.length > 0 ? Number((vendidasPeriodo.length / cotizacionesPeriodo.length * 100).toFixed(1)) : 0;
    const clientesCotizados = new Set(cotizacionesPeriodo.map(c => String(c.cliente || "").trim()).filter(Boolean));
    const clientesVendidos = new Set(ventasPeriodo.map(n => String(n.cliente || "").trim()).filter(Boolean));
    const ventasPorCliente = ventasPeriodo.reduce((acc, n) => {
      const cliente = String(n.cliente || "Sin cliente").trim() || "Sin cliente";
      if (!acc[cliente]) acc[cliente] = { cliente, cantidad:0, total:0 };
      acc[cliente].cantidad += 1;
      acc[cliente].total += Number(n.total || 0);
      return acc;
    }, {});
    const topClientes = Object.values(ventasPorCliente).sort((a,b) => b.total - a.total).slice(0, 10);

    const montoCotizadoVendido = montoVendido;
    const conversionMonto = montoCotizado > 0 ? Number((montoVendido / montoCotizado * 100).toFixed(1)) : 0;

    const sumarEstado = (estado) => cotizacionesConEstado
      .filter(q => q.status === estado)
      .reduce((s,q) => s + Number(q.total || 0), 0);

    const embudoDinero = [
      { key:"vendida", label:"Vendidas", cantidad:vendidasPeriodo.length, monto:sumarEstado("vendida"), color:COLORS.success },
      { key:"activa", label:"Activas", cantidad:activasPeriodo.length, monto:sumarEstado("activa"), color:"#5a8abe" },
      { key:"urgente", label:"Seguimiento", cantidad:urgentesPeriodo.length, monto:sumarEstado("urgente"), color:COLORS.warning },
      { key:"vencida", label:"Vencidas", cantidad:vencidasPeriodo.length, monto:sumarEstado("vencida"), color:"#9a7aaa" },
    ];

    const tramosBase = [
      { key:"hasta_100", label:"Hasta $100.000", min:0, max:100000 },
      { key:"100_250", label:"$100.001 - $250.000", min:100001, max:250000 },
      { key:"250_500", label:"$250.001 - $500.000", min:250001, max:500000 },
      { key:"500_1m", label:"$500.001 - $1.000.000", min:500001, max:1000000 },
      { key:"mas_1m", label:"Más de $1.000.000", min:1000001, max:Infinity },
    ];

    const conversionPorTramo = tramosBase.map((t) => {
      const items = cotizacionesConEstado.filter(q => {
        const total = Number(q.total || 0);
        return total >= t.min && total <= t.max;
      });
      const vendidasTramo = items.filter(q => q.status === "vendida");
      const montoTramo = items.reduce((s,q) => s + Number(q.total || 0), 0);
      const montoVendidoTramo = vendidasTramo.reduce((s,q) => s + Number(q.total || 0), 0);
      return {
        ...t,
        cantidad:items.length,
        vendidas:vendidasTramo.length,
        conversion:items.length ? Number((vendidasTramo.length / items.length * 100).toFixed(1)) : 0,
        monto:montoTramo,
        montoVendido:montoVendidoTramo,
        conversionMonto:montoTramo ? Number((montoVendidoTramo / montoTramo * 100).toFixed(1)) : 0,
      };
    });

    const cierres = vendidasPeriodo.map((q) => {
      const fechaCot = obtenerFechaValida(q.fecha);
      const nv = notasPeriodo
        .filter(n => String(n.cotizacion) === String(q.numero))
        .map(n => ({ ...n, fechaObj: obtenerFechaValida(n.fecha) }))
        .filter(n => n.fechaObj && fechaCot)
        .sort((a,b) => a.fechaObj - b.fechaObj)[0];
      if (!fechaCot || !nv?.fechaObj) return null;
      return Math.max(0, Math.round((nv.fechaObj - fechaCot) / (1000 * 60 * 60 * 24)));
    }).filter(v => v !== null);

    const diasPromedioCierre = cierres.length
      ? Number((cierres.reduce((s,d) => s + d, 0) / cierres.length).toFixed(1))
      : 0;

    return {
      cotizacionesPeriodo,
      notasPeriodo,
      barranesPeriodo,
      ventasPeriodo,
      cotizacionesConEstado,
      vendidasPeriodo,
      activasPeriodo,
      urgentesPeriodo,
      vencidasPeriodo,
      montoCotizado,
      montoVendido,
      conversion,
      conversionMonto,
      montoCotizadoVendido,
      embudoDinero,
      conversionPorTramo,
      diasPromedioCierre,
      ticketCotizado: cotizacionesPeriodo.length ? Math.round(montoCotizado / cotizacionesPeriodo.length) : 0,
      ticketVendido: ventasPeriodo.length ? Math.round(montoVendido / ventasPeriodo.length) : 0,
      clientesCotizados: clientesCotizados.size,
      clientesVendidos: clientesVendidos.size,
      topClientes
    };
  };

  const rangoResumen = obtenerRangoResumen();
  const resumenActual = calcularMetricasResumen(rangoResumen);
  const rangoResumenAnterior = filtroResumen === "todo" ? null : obtenerRangoResumen(filtroResumen, filtroResumen === "año_actual" ? -12 : filtroResumen === "ultimos_6_meses" ? -6 : filtroResumen === "ultimos_3_meses" ? -3 : -1);
  const resumenAnterior = rangoResumenAnterior ? calcularMetricasResumen(rangoResumenAnterior) : null;
  const variacionPorcentual = (actual, anterior) => {
    if (!anterior) return actual > 0 ? 100 : 0;
    return Number(((actual - anterior) / anterior * 100).toFixed(1));
  };
  const diferenciaConversionResumen = resumenAnterior ? Number((resumenActual.conversion - resumenAnterior.conversion).toFixed(1)) : 0;
  const variacionVentasResumen = resumenAnterior ? variacionPorcentual(resumenActual.montoVendido, resumenAnterior.montoVendido) : 0;
  const mesesResumenGrafico = (() => {
    const hoy = new Date();
    const cantidad = filtroResumen === "ultimos_6_meses" ? 6 : filtroResumen === "ultimos_3_meses" ? 3 : 6;
    return Array.from({ length:cantidad }, (_, i) => {
      const fecha = sumarMeses(hoy, i - cantidad + 1);
      const key = `${fecha.getFullYear()}-${String(fecha.getMonth()+1).padStart(2,"0")}`;
      const rangoMes = { desde:inicioMes(fecha), hasta:finMes(fecha) };
      const m = calcularMetricasResumen(rangoMes);
      return { key, label:nombreMesResumen(key).replace(` ${fecha.getFullYear()}`, ""), ...m };
    });
  })();
  const maxGraficoResumen = Math.max(...mesesResumenGrafico.map(m => Math.max(m.montoCotizado, m.montoVendido)), 1);

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

const obtenerNumeroCotizacionOrden = (valor) => {
  const numeroLimpio = String(valor || "").replace(/\D/g, "");
  return Number(numeroLimpio || 0);
};

const filteredQuotes = withStatus.filter(q =>
  (String(q.numero || "").toLowerCase().includes(filter.toLowerCase()) || String(q.cliente || "").toLowerCase().includes(filter.toLowerCase())) &&
  cumpleFecha(q.fecha) &&
  cumpleMes(q.fecha) &&
  (showVencidas || q.status !== "vencida")
).sort((a,b) => {
  return obtenerNumeroCotizacionOrden(b.numero) - obtenerNumeroCotizacionOrden(a.numero);
});

const filteredNotas = notasVenta.filter(s =>
  (s.numero.toLowerCase().includes(filter.toLowerCase()) || s.cliente.toLowerCase().includes(filter.toLowerCase())) &&
  cumpleFecha(s.fecha) &&
  cumpleMes(s.fecha)
);

const filteredBarranes = barranes.filter(s =>
  (s.numero.toLowerCase().includes(filter.toLowerCase()) || s.cliente.toLowerCase().includes(filter.toLowerCase())) &&
  cumpleFecha(s.fecha) &&
  cumpleMes(s.fecha)
);

const totalSoldFiltrado = filteredNotas.reduce((s,n) => s + Number(n.total || 0), 0);
const totalBarranesFiltrado = filteredBarranes.reduce((s,n) => s + Number(n.total || 0), 0);
const mesSeleccionadoTexto = nombreMes(mesFiltro);

  const r2=58,cx=80,cy=76;
  const toRad=d=>d*Math.PI/180;
  const startA=200,sweepA=140;
  const arcPath=(s,e)=>{
    const p1={x:cx+r2*Math.cos(toRad(s)),y:cy+r2*Math.sin(toRad(s))};
    const p2={x:cx+r2*Math.cos(toRad(e)),y:cy+r2*Math.sin(toRad(e))};
    return `M ${p1.x} ${p1.y} A ${r2} ${r2} 0 ${e-s>180?1:0} 1 ${p2.x} ${p2.y}`;
  };
  const conteoCotizacionesEnNV = [...notas, ...barranes].reduce((acc, nv) => {
  const cot = String(nv.cotizacion || "").trim();

  if (!cot) return acc;

  acc[cot] = (acc[cot] || 0) + 1;

  return acc;
}, {});
  const fillEnd=startA+sweepA*(rate/100);

  const guardarEdicionCompletaCotizacion = async ({ cotizacionOriginal, cotizacionEditada, detallesEditados }) => {
    const idReal = Number(String(cotizacionOriginal.id).replace("supabase-", ""));
    const numeroAnterior = String(cotizacionOriginal.numero || "");
    const numeroNuevo = String(cotizacionEditada.numero || "").trim();

    const { error } = await supabase
      .from("cotizaciones")
      .update({
        numero: Number(numeroNuevo) || numeroNuevo,
        cliente: cotizacionEditada.cliente,
        fecha_creacion: cotizacionEditada.fecha,
        total: Number(cotizacionEditada.total || 0)
      })
      .eq("id", idReal);

    if (error) {
      console.error(error);
      toast("No se pudo actualizar la cotización.");
      return;
    }

    await supabase
      .from("detalles_cotizaciones")
      .delete()
      .eq("cotizacion_id", idReal);

    let detallesGuardados = [];

    if (detallesEditados.length > 0) {
      const detallesParaGuardar = detallesEditados.map((d) => ({
        cotizacion_id: idReal,
        unidad: d.unidad,
        tipo: d.tipo,
        largo: d.largo,
        ancho: d.ancho,
        color: d.color,
        valor: d.valor,
        total: d.total
      }));

      const { data, error: errorInsert } = await supabase
        .from("detalles_cotizaciones")
        .insert(detallesParaGuardar)
        .select("*");

      if (errorInsert) {
        console.error(errorInsert);
        toast("La cotización se actualizó, pero hubo un error guardando el detalle.");
        return;
      }

      detallesGuardados = data || detallesParaGuardar;
    }

    setCotizaciones(prev => prev.map(c =>
      c.id === cotizacionOriginal.id
        ? { ...c, ...cotizacionEditada, numero: numeroNuevo, total: Number(cotizacionEditada.total || 0) }
        : c
    ));

    setNotas(prev => prev.map(n =>
      String(n.cotizacion) === numeroAnterior
        ? { ...n, cotizacion: numeroNuevo }
        : n
    ));

    setDetallesCotizaciones(prev => [
      ...prev.filter(d => Number(d.cotizacion_id) !== idReal),
      ...detallesGuardados
    ]);

    if (modalCot?.id === cotizacionOriginal.id) {
      setModalCot({ ...modalCot, ...cotizacionEditada, numero: numeroNuevo, total: Number(cotizacionEditada.total || 0) });
    }

    setModalEditarCotizacion(null);
    toast("Cotización actualizada correctamente.");
  };

  const guardarEdicionCompletaNV = async ({ notaOriginal, notaEditada, detallesEditados }) => {
    const idReal = Number(String(notaOriginal.id).replace("supabase-", ""));
    const numeroAnterior = String(notaOriginal.numero || "");
    const numeroNuevo = String(notaEditada.numero || "").trim();
    const cotizacionNumero = String(notaEditada.cotizacion || "").trim();

    let cotizacionId = null;

    if (cotizacionNumero) {
      const { data: cotizacionEncontrada, error: errorCotizacion } = await supabase
        .from("cotizaciones")
        .select("id")
        .eq("numero", cotizacionNumero)
        .maybeSingle();

      if (errorCotizacion) {
        console.error(errorCotizacion);
      }

      if (cotizacionEncontrada?.id) {
        cotizacionId = cotizacionEncontrada.id;
      }
    }

    const { error } = await supabase
      .from("notas_venta")
      .update({
        numero: Number(numeroNuevo) || numeroNuevo,
        cliente: notaEditada.cliente,
        fecha: notaEditada.fecha,
        total: Number(notaEditada.total || 0),
        cotizacion_id: cotizacionId
      })
      .eq("id", idReal);

    if (error) {
      console.error(error);
      toast("No se pudo actualizar la Nota de Venta.");
      return;
    }

    const { error: errorDeleteAnterior } = await supabase
      .from("detalles_notas_venta_produccion")
      .delete()
      .eq("nota_venta_numero", numeroAnterior);

    if (errorDeleteAnterior) {
      console.error(errorDeleteAnterior);
      toast("La NV se actualizó, pero no se pudo reemplazar el detalle anterior.");
      return;
    }

    if (numeroNuevo !== numeroAnterior) {
      await supabase
        .from("detalles_notas_venta_produccion")
        .delete()
        .eq("nota_venta_numero", numeroNuevo);
    }

    const detallesParaGuardar = detallesEditados.map((d, idx) => ({
      nota_venta_numero: numeroNuevo,
      cotizacion_numero: cotizacionNumero,
      cliente: notaEditada.cliente,
      material: d.material,
      cantidad: d.cantidad,
      descripcion: d.descripcion,
      alto: d.alto,
      ancho: d.ancho,
      color: d.color,
      orden: idx + 1
    }));

    let detallesGuardados = [];

    if (detallesParaGuardar.length > 0) {
      const { data, error: errorInsert } = await supabase
        .from("detalles_notas_venta_produccion")
        .insert(detallesParaGuardar)
        .select("*");

      if (errorInsert) {
        console.error(errorInsert);
        toast("La NV se actualizó, pero hubo un error guardando el nuevo detalle.");
        return;
      }

      detallesGuardados = data || detallesParaGuardar;
    }

    setNotas(prev => prev.map(n =>
      n.id === notaOriginal.id
        ? { ...n, ...notaEditada, numero: numeroNuevo, cotizacion: cotizacionNumero || null, total: Number(notaEditada.total || 0) }
        : n
    ));

    setDetallesNotasVenta(prev => [
      ...prev.filter(d => String(d.nota_venta_numero) !== numeroAnterior && String(d.nota_venta_numero) !== numeroNuevo),
      ...detallesGuardados
    ]);

    if (modalNV?.id === notaOriginal.id) {
      setModalNV({ ...modalNV, ...notaEditada, numero: numeroNuevo, cotizacion: cotizacionNumero || null, total: Number(notaEditada.total || 0) });
    }

    setModalEditarNV(null);
    toast("Nota de Venta actualizada correctamente.");
  };

  const guardarGestionNV = async (nvActualizada) => {
  const idReal = Number(String(nvActualizada.id).replace("supabase-", ""));

  const nuevoAbono = Number(nvActualizada.nuevoAbono) || 0;
  const fechaAbono = nvActualizada.fechaAbono || new Date().toISOString().split("T")[0];
  const medioPagoAbono = nvActualizada.medioPagoAbono || "transferencia";

  if (nuevoAbono > 0) {
    const { error: errorInsert } = await supabase
      .from("abonos_nv")
      .insert({
        nota_venta_id: idReal,
        monto: nuevoAbono,
        fecha: fechaAbono,
        medio_pago: medioPagoAbono,
        observacion: nvActualizada.observacionAbono || ""
      });

    if (errorInsert) {
      toast("Error al guardar el abono");
      console.error(errorInsert);
      return;
    }
  }

  const { data: abonosActualizados, error: errorAbonos } = await supabase
    .from("abonos_nv")
    .select("*")
    .eq("nota_venta_id", idReal);

  if (errorAbonos) {
    toast("Error al calcular abonos");
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
    toast("Error al guardar gestión de NV");
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
      fecha: fechaAbono,
      numero: nvActualizada.numero,
      cliente: nvActualizada.cliente,
      monto: nuevoAbono,
      medioPago: medioPagoAbono,
      tipoDocumento: nvActualizada.tipo_documento || "nv"
    });
  }

  setModalNV(null);
};
  
const marcarComoEntregadaProduccion = async (nota) => {
  const confirmar = await confirmDialog(`¿Marcar ${((nota.tipo_documento || "nv") === "barran" ? "Barrán" : "NV")} ${nota.numero} como entregada?\n\nDesaparecerá del listado de Producción.`);
  if (!confirmar) return;

  const idReal = Number(String(nota.id).replace("supabase-", ""));

  const { error } = await supabase
    .from("notas_venta")
    .update({ proceso: "entregado" })
    .eq("id", idReal);

  if (error) {
    console.error(error);
    toast("No se pudo marcar como entregada.");
    return;
  }

  setNotas(prev => prev.map(n =>
    n.id === nota.id ? { ...n, proceso: "entregado" } : n
  ));

  setModalProduccion(null);
};

const guardarProduccion = async (notaActualizada) => {
  const idReal = Number(String(notaActualizada.id).replace("supabase-", ""));
  const procesoAutomatico = calcularProcesoDesdeAvance(notaActualizada);

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
      media_cana: notaActualizada.media_cana,
      proceso: procesoAutomatico,
    })
    .eq("id", idReal);

  if (error) {
    toast("Error al guardar producción");
    console.error(error);
    return;
  }

  setNotas(notas.map(n =>
    n.id === notaActualizada.id
      ? { ...n, ...notaActualizada, proceso: procesoAutomatico }
      : n
  ));

  setModalProduccion(null);
};

const guardarCuentaPorPagar = async (e) => {
  e.preventDefault();
  const form = new FormData(e.currentTarget);
  const montoTotal = Number(form.get("monto_total") || 0);
  const montoCuota = Number(form.get("monto_cuota") || 0);
  const payload = {
    nombre: String(form.get("nombre") || "").trim(),
    tipo: String(form.get("tipo") || "proveedor"),
    detalle: String(form.get("detalle") || "").trim(),
    monto_total: montoTotal,
    monto_cuota: montoCuota || montoTotal,
    fecha_vencimiento: String(form.get("fecha_vencimiento") || new Date().toISOString().split("T")[0]),
    cuota_actual: Number(form.get("cuota_actual") || 1),
    cuotas_totales: Number(form.get("cuotas_totales") || 1),
    estado: "pendiente"
  };

  if (!payload.nombre || !payload.monto_total) {
    toast("Debes ingresar nombre y monto.");
    return;
  }

  const { data, error } = await supabase
    .from("cuentas_pagar")
    .insert(payload)
    .select("*")
    .single();

  if (error) {
    console.error(error);
    toast("No se pudo guardar la cuenta por pagar.");
    return;
  }

  setCuentasPagar(prev => [data, ...prev]);
  setModalCuentaPagar(false);
};

const guardarAbonoCuentaPagar = async (e) => {
  e.preventDefault();
  if (!modalAbonoCuentaPagar) return;
  const form = new FormData(e.currentTarget);
  const monto = Number(form.get("monto") || 0);
  const fecha = String(form.get("fecha") || new Date().toISOString().split("T")[0]);
  const medioPago = String(form.get("medio_pago") || "transferencia");
  const observacion = String(form.get("observacion") || "").trim();

  if (!monto) {
    toast("Debes ingresar un monto de abono.");
    return;
  }

  const { data, error } = await supabase
    .from("cuentas_pagar_abonos")
    .insert({ cuenta_pagar_id: modalAbonoCuentaPagar.id, monto, fecha, medio_pago: medioPago, observacion })
    .select("*")
    .single();

  if (error) {
    console.error(error);
    toast("No se pudo guardar el abono.");
    return;
  }

  const abonosDeCuenta = [...abonosCuentasPagar.filter(a => a.cuenta_pagar_id === modalAbonoCuentaPagar.id), data];
  const totalAbonado = abonosDeCuenta.reduce((sum, a) => sum + Number(a.monto || 0), 0);
  const estado = totalAbonado >= Number(modalAbonoCuentaPagar.monto_total || 0) ? "pagado" : "pendiente";

  await supabase
    .from("cuentas_pagar")
    .update({ estado })
    .eq("id", modalAbonoCuentaPagar.id);

  const cuentaHaber = medioPago === "efectivo" ? "Caja" : medioPago === "cheque" ? "Cheques por pagar" : "Banco";
  const documentoPagoCxp = `PAGO-CXP-${modalAbonoCuentaPagar.id}-${Date.now()}`;
  await guardarAsientosAutomaticos([
    {
      fecha,
      detalle: `Abono cuenta por pagar - ${modalAbonoCuentaPagar.nombre}`,
      desglose: "Pago cuenta por pagar",
      documento: documentoPagoCxp,
      definicion: "Proveedores",
      debe: monto,
      haber: 0
    },
    {
      fecha,
      detalle: `Abono cuenta por pagar - ${modalAbonoCuentaPagar.nombre}`,
      desglose: "Pago cuenta por pagar",
      documento: documentoPagoCxp,
      definicion: cuentaHaber,
      debe: 0,
      haber: monto
    }
  ]);

  setAbonosCuentasPagar(prev => [data, ...prev]);
  setCuentasPagar(prev => prev.map(c => c.id === modalAbonoCuentaPagar.id ? { ...c, estado } : c));
  setModalAbonoCuentaPagar(null);
};


const limpiarNumeroCartola = (valor) => {
  if (valor === null || valor === undefined || valor === "") return 0;
  if (typeof valor === "number") return Math.round(valor);
  const texto = String(valor)
    .replace(/\$/g, "")
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(/,/g, ".")
    .replace(/[^0-9.-]/g, "");
  const numero = Number(texto || 0);
  return Number.isFinite(numero) ? Math.round(numero) : 0;
};

const normalizarTextoCartola = (valor) => String(valor || "")
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "")
  .toLowerCase()
  .trim();

const convertirFechaCartola = (valor) => {
  if (!valor) return new Date().toISOString().split("T")[0];
  if (typeof valor === "number") {
    const parsed = XLSX.SSF.parse_date_code(valor);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  const texto = String(valor).trim();
  const iso = texto.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return `${iso[1]}-${String(iso[2]).padStart(2, "0")}-${String(iso[3]).padStart(2, "0")}`;
  const chilena = texto.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (chilena) {
    const year = chilena[3].length === 2 ? `20${chilena[3]}` : chilena[3];
    return `${year}-${String(chilena[2]).padStart(2, "0")}-${String(chilena[1]).padStart(2, "0")}`;
  }
  return new Date().toISOString().split("T")[0];
};

const clasificarMovimientoCartola = ({ descripcion, cargo, abono }) => {
  const desc = normalizarTextoCartola(descripcion);
  if (abono > 0) {
    if (desc.includes("deposito") || desc.includes("transferencia") || desc.includes("abono")) return "Ingreso por revisar";
    return "Ingreso por revisar";
  }
  if (desc.includes("previred") || desc.includes("afp") || desc.includes("fonasa") || desc.includes("isapre")) return "Imposiciones";
  if (desc.includes("sueldo") || desc.includes("remuneracion") || desc.includes("anticipo")) return "Remuneraciones";
  if (desc.includes("arriendo")) return "Arriendo";
  if (desc.includes("bencina") || desc.includes("combustible") || desc.includes("copec") || desc.includes("shell") || desc.includes("petrobras")) return "Combustible";
  if (desc.includes("luz") || desc.includes("agua") || desc.includes("internet") || desc.includes("entel") || desc.includes("wom") || desc.includes("claro") || desc.includes("movistar")) return "Servicios";
  if (desc.includes("cheque")) return "Cheque por revisar";
  if (cargo > 0) return "Egreso por revisar";
  return "Por revisar";
};

const importarCartolaBancaria = async (event) => {
  const files = Array.from(event.target.files || []);
  event.target.value = "";
  if (files.length === 0) return;

  const movimientosDetectados = [];

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type:"array", cellDates:false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
    if (!rows.length) continue;

    let headerIndex = rows.findIndex(row => {
      const normalizadas = row.map(normalizarTextoCartola);
      return normalizadas.some(c => c.includes("fecha")) &&
        normalizadas.some(c => c.includes("descripcion") || c.includes("detalle") || c.includes("glosa") || c.includes("movimiento"));
    });

    if (headerIndex < 0) headerIndex = 0;

    const headers = rows[headerIndex].map(normalizarTextoCartola);
    const buscarCol = (opciones) => headers.findIndex(h => opciones.some(op => h.includes(op)));

    const colFecha = buscarCol(["fecha"]);
    const colDescripcion = buscarCol(["descripcion", "detalle", "glosa", "movimiento"]);
    let colCargo = buscarCol(["cargo", "egreso", "retiro", "debe", "salida"]);
    let colAbono = buscarCol(["abono", "ingreso", "deposito", "haber", "entrada"]);
    const colMonto = buscarCol(["monto", "importe", "valor"]);
    const colSaldo = buscarCol(["saldo"]);

    for (const row of rows.slice(headerIndex + 1)) {
      const fecha = convertirFechaCartola(colFecha >= 0 ? row[colFecha] : row[0]);
      const descripcion = String(colDescripcion >= 0 ? row[colDescripcion] : row[1] || "").trim();
      if (!descripcion) continue;

      let cargo = colCargo >= 0 ? Math.abs(limpiarNumeroCartola(row[colCargo])) : 0;
      let abono = colAbono >= 0 ? Math.abs(limpiarNumeroCartola(row[colAbono])) : 0;

      if (cargo === 0 && abono === 0 && colMonto >= 0) {
        const monto = limpiarNumeroCartola(row[colMonto]);
        if (monto < 0) cargo = Math.abs(monto);
        if (monto > 0) abono = Math.abs(monto);
      }

      if (cargo === 0 && abono === 0) continue;

      const saldo = colSaldo >= 0 ? limpiarNumeroCartola(row[colSaldo]) : null;
      const categoria = clasificarMovimientoCartola({ descripcion, cargo, abono });

      movimientosDetectados.push({
        fecha,
        descripcion,
        cargo,
        abono,
        saldo,
        categoria,
        estado:"por_revisar",
        archivo_origen:file.name
      });
    }
  }

  if (movimientosDetectados.length === 0) {
    toast("No se detectaron movimientos en la cartola. Revisa que el archivo tenga fecha, descripción y cargos/abonos.");
    return;
  }

  const { data, error } = await supabase
    .from("cartola_bancaria")
    .insert(movimientosDetectados)
    .select("*");

  if (error) {
    console.error(error);
    toast("No se pudo guardar la cartola. Revisa si ejecutaste el SQL de cartola_bancaria.");
    return;
  }

  setCartolaMovimientos(prev => [...(data || []), ...prev]);
  setMesCartola(movimientosDetectados[0]?.fecha?.slice(0, 7) || mesCartola);
  toast(`Cartola importada: ${movimientosDetectados.length} movimientos detectados.`);
};

const actualizarMovimientoCartola = async (movimiento, cambios) => {
  const { data, error } = await supabase
    .from("cartola_bancaria")
    .update(cambios)
    .eq("id", movimiento.id)
    .select("*")
    .single();

  if (error) {
    console.error(error);
    toast("No se pudo actualizar el movimiento de cartola.");
    return null;
  }

  setCartolaMovimientos(prev => prev.map(m => m.id === movimiento.id ? data : m));
  return data;
};

const registrarAbonoDocumento = async ({
  nota,
  monto,
  fecha,
  medioPago = "transferencia",
  tipoAbono = "abono",
  observacion = "",
  movimiento = null
}) => {
  if (!nota) {
    toast("Selecciona una NV o Barrán para registrar el abono.", "warning");
    return false;
  }

  const montoBase = Math.round(Number(monto || 0));
  if (!montoBase || montoBase <= 0) {
    toast("El monto del abono debe ser mayor a cero.", "warning");
    return false;
  }

  const idReal = Number(String(nota.id).replace("supabase-", ""));
  const abonosExistentes = abonosNV.filter(a => Number(a.nota_venta_id) === idReal);
  const abonadoAntes = abonosExistentes.reduce((sum, a) => sum + Number(a.monto || 0), 0);
  const totalDocumento = Math.round(Number(nota.total || 0));
  const saldoAntes = Math.max(totalDocumento - abonadoAntes, 0);

  const etiquetasTipoAbono = {
    abono_inicial: "Abono inicial",
    abono: "Abono parcial",
    abono_retiro_parcial: "Abono retiro parcial",
    abono_saldo: "Abono saldo"
  };

  let registros = [{
    monto: montoBase,
    fecha: fecha || new Date().toISOString().split("T")[0],
    medio_pago: medioPago,
    observacion: `${etiquetasTipoAbono[tipoAbono] || "Abono"}${observacion ? ` · ${observacion}` : ""}`.trim()
  }];

  const saldoDespues = Math.max(saldoAntes - montoBase, 0);

  if (saldoDespues > 0 && saldoDespues <= 99) {
    const cerrar = await confirmDialog(
      `Queda un saldo de ${fmt(saldoDespues)}.\n\n¿Deseas cerrar la venta con ajuste comercial?`
    );

    if (cerrar) {
      registros.push({
        monto: saldoDespues,
        fecha: fecha || new Date().toISOString().split("T")[0],
        medio_pago: "ajuste",
        observacion: `Ajuste comercial automático por diferencia menor o igual a $99. Saldo ajustado: ${fmt(saldoDespues)}`
      });
    }
  } else if (tipoAbono === "abono_saldo" && saldoDespues > 99) {
    const confirmar = await confirmDialog(
      `Este abono no cubre todo el saldo.\n\nSaldo pendiente después del abono: ${fmt(saldoDespues)}.\n\n¿Deseas registrarlo igualmente como abono parcial?`
    );
    if (!confirmar) return false;
    registros = registros.map(r => ({
      ...r,
      observacion: r.observacion.replace("Abono saldo", "Abono parcial")
    }));
  }

  const payload = registros.map(r => ({
    nota_venta_id: idReal,
    monto: r.monto,
    fecha: r.fecha,
    medio_pago: r.medio_pago,
    observacion: r.observacion
  }));

  const { error: errorInsert } = await supabase
    .from("abonos_nv")
    .insert(payload);

  if (errorInsert) {
    console.error(errorInsert);
    toast("No se pudo registrar el abono.", "error");
    return false;
  }

  const { data: abonosActualizados, error: errorAbonos } = await supabase
    .from("abonos_nv")
    .select("*")
    .eq("nota_venta_id", idReal);

  if (errorAbonos) {
    console.error(errorAbonos);
  }

  const totalAbonado = (abonosActualizados || [...abonosExistentes, ...payload])
    .reduce((sum, a) => sum + Number(a.monto || 0), 0);

  const estadoPago = totalAbonado >= totalDocumento && totalDocumento > 0
    ? "pagada"
    : totalAbonado > 0
    ? "abonada"
    : "pendiente";

  await supabase
    .from("notas_venta")
    .update({ abono: totalAbonado, estado_pago: estadoPago })
    .eq("id", idReal);

  await crearAsientoAbonoNVAutomatico({
    fecha: fecha || new Date().toISOString().split("T")[0],
    numero: nota.numero,
    cliente: nota.cliente,
    monto: montoBase,
    medioPago,
    tipoDocumento: nota.tipo_documento || "nv"
  });

  setAbonosNV(prev => [...prev.filter(a => Number(a.nota_venta_id) !== idReal), ...(abonosActualizados || payload)]);
  setNotas(prev => prev.map(n => n.id === nota.id ? { ...n, abono: totalAbonado, estado_pago: estadoPago } : n));

  if (movimiento) {
    await actualizarMovimientoCartola(movimiento, {
      estado: "conciliado",
      categoria: "Abono CxC",
      nota_venta_id: idReal
    });
  }

  toast(`Abono registrado en ${(nota.tipo_documento || "nv") === "barran" ? "Barrán" : "NV"} #${nota.numero}.`, "success");
  return true;
};

const registrarCartolaComoAbonoNV = async (movimiento) => {
  setModalClasificarCartola(movimiento);
  setCategoriaCartolaSeleccionada("Abono CxC");
  setBusquedaCategoriaCartola("");
  setBusquedaCxcCartola("");
  setCxcSeleccionadaCartola(null);
  setTipoAbonoCartola("abono");
};

const abrirClasificacionCartola = (movimiento) => {
  setModalClasificarCartola(movimiento);
  setCategoriaCartolaSeleccionada("");
  setBusquedaCategoriaCartola("");
  setBusquedaCxcCartola("");
  setCxcSeleccionadaCartola(null);
  setTipoAbonoCartola(Number(movimiento?.abono || 0) > 0 ? "abono" : "");
};

const cerrarClasificacionCartola = () => {
  setModalClasificarCartola(null);
  setCategoriaCartolaSeleccionada("");
  setBusquedaCategoriaCartola("");
  setBusquedaCxcCartola("");
  setCxcSeleccionadaCartola(null);
  setTipoAbonoCartola("abono");
};

const guardarClasificacionCartola = async () => {
  const movimiento = modalClasificarCartola;
  if (!movimiento) return;

  const categoria = String(categoriaCartolaSeleccionada || "").trim();
  if (!categoria) {
    toast("Selecciona una categoría.", "warning");
    return;
  }

  if (categoria === "Abono CxC") {
    const monto = Number(movimiento.abono || 0);
    if (!monto) {
      toast("Este movimiento no tiene abono/ingreso para asociar a una CxC.", "warning");
      return;
    }

    if (!cxcSeleccionadaCartola) {
      toast("Selecciona la NV o Barrán al que corresponde el abono.", "warning");
      return;
    }

    const ok = await registrarAbonoDocumento({
      nota: cxcSeleccionadaCartola.nota,
      monto,
      fecha: movimiento.fecha,
      medioPago: "transferencia",
      tipoAbono: tipoAbonoCartola || "abono",
      observacion: `Registrado desde cartola bancaria. ${movimiento.descripcion || ""}`,
      movimiento
    });

    if (ok) cerrarClasificacionCartola();
    return;
  }

  await actualizarMovimientoCartola(movimiento, {
    estado: "conciliado",
    categoria
  });

  toast(`Movimiento clasificado como ${categoria}.`, "success");
  cerrarClasificacionCartola();
};

const normalizarClienteCredito = (cliente) =>
  String(cliente || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();

const clienteTieneCreditoAutorizado = (cliente) => {
  const clave = normalizarClienteCredito(cliente);
  return clientesCreditoAutorizado.some(c => normalizarClienteCredito(c) === clave);
};

const toggleCreditoAutorizadoCliente = async (cliente) => {
  const nombre = String(cliente || "").trim();
  if (!nombre) return;

  const yaExiste = clienteTieneCreditoAutorizado(nombre);
  const nuevos = yaExiste
    ? clientesCreditoAutorizado.filter(c => normalizarClienteCredito(c) !== normalizarClienteCredito(nombre))
    : [...clientesCreditoAutorizado, nombre];

  setClientesCreditoAutorizado(nuevos);
  try {
    localStorage.setItem("sf-clientes-credito-autorizado", JSON.stringify(nuevos));
  } catch (e) {}

  toast(
    yaExiste
      ? `${nombre} ya no está marcado con crédito autorizado.`
      : `${nombre} quedó marcado con crédito autorizado.`,
    yaExiste ? "info" : "success"
  );
};

const registrarCartolaComoPagoCuenta = async (movimiento) => {
  const monto = Number(movimiento.cargo || 0);
  if (!monto) {
    toast("Este movimiento no tiene cargo/egreso para asociar a una cuenta por pagar.");
    return;
  }

  const texto = window.prompt("Escribe el nombre o ID de la cuenta por pagar:");
  if (!texto) return;

  const busqueda = normalizarTextoCartola(texto);
  const cuenta = cuentasPagar.find(c => String(c.id) === String(texto).trim()) ||
    cuentasPagar.find(c => normalizarTextoCartola(c.nombre).includes(busqueda));

  if (!cuenta) {
    toast("No encontré esa cuenta por pagar.");
    return;
  }

  const { data, error } = await supabase
    .from("cuentas_pagar_abonos")
    .insert({ cuenta_pagar_id:cuenta.id, monto, fecha:movimiento.fecha, medio_pago:"transferencia", observacion:`Pago desde cartola. ${movimiento.descripcion || ""}` })
    .select("*")
    .single();

  if (error) {
    console.error(error);
    toast("No se pudo guardar el abono de la cuenta por pagar.");
    return;
  }

  const abonosDeCuenta = [...abonosCuentasPagar.filter(a => String(a.cuenta_pagar_id) === String(cuenta.id)), data];
  const totalAbonado = abonosDeCuenta.reduce((sum, a) => sum + Number(a.monto || 0), 0);
  const estado = totalAbonado >= Number(cuenta.monto_total || 0) ? "pagado" : "pendiente";

  await supabase.from("cuentas_pagar").update({ estado }).eq("id", cuenta.id);

  const documentoPagoCxp = `CARTOLA-CXP-${movimiento.id}-${Date.now()}`;
  await guardarAsientosAutomaticos([
    { fecha:movimiento.fecha, detalle:`Pago cuenta por pagar - ${cuenta.nombre}`, desglose:"Pago desde cartola bancaria", documento:documentoPagoCxp, definicion:"Proveedores", debe:monto, haber:0 },
    { fecha:movimiento.fecha, detalle:`Pago cuenta por pagar - ${cuenta.nombre}`, desglose:"Pago desde cartola bancaria", documento:documentoPagoCxp, definicion:"Banco", debe:0, haber:monto }
  ]);

  setAbonosCuentasPagar(prev => [data, ...prev]);
  setCuentasPagar(prev => prev.map(c => c.id === cuenta.id ? { ...c, estado } : c));
  await actualizarMovimientoCartola(movimiento, { estado:"conciliado", categoria:"Pago cuenta por pagar", cuenta_pagar_id:cuenta.id });
};

const registrarCartolaComoMovimientoContable = async (movimiento) => {
  const montoCargo = Number(movimiento.cargo || 0);
  const montoAbono = Number(movimiento.abono || 0);
  const cuenta = window.prompt("Cuenta contable para este movimiento:", movimiento.categoria || (montoCargo > 0 ? "Gastos generales" : "Ingresos por revisar"));
  if (!cuenta) return;

  const documento = `CARTOLA-${movimiento.id}-${Date.now()}`;

  if (montoCargo > 0) {
    await guardarAsientosAutomaticos([
      { fecha:movimiento.fecha, detalle:movimiento.descripcion, desglose:"Cartola bancaria", documento, definicion:cuenta, debe:montoCargo, haber:0 },
      { fecha:movimiento.fecha, detalle:movimiento.descripcion, desglose:"Cartola bancaria", documento, definicion:"Banco", debe:0, haber:montoCargo }
    ]);
  } else if (montoAbono > 0) {
    await guardarAsientosAutomaticos([
      { fecha:movimiento.fecha, detalle:movimiento.descripcion, desglose:"Cartola bancaria", documento, definicion:"Banco", debe:montoAbono, haber:0 },
      { fecha:movimiento.fecha, detalle:movimiento.descripcion, desglose:"Cartola bancaria", documento, definicion:cuenta, debe:0, haber:montoAbono }
    ]);
  }

  await actualizarMovimientoCartola(movimiento, { estado:"conciliado", categoria:cuenta, asiento_documento:documento });
};

const guardarTrabajador = async (e) => {
  e.preventDefault();
  const form = new FormData(e.currentTarget);
  const payload = {
    nombre: String(form.get("nombre") || "").trim(),
    rut: String(form.get("rut") || "").trim(),
    cargo: String(form.get("cargo") || "").trim(),
    fecha_ingreso: String(form.get("fecha_ingreso") || "") || null,
    sueldo_base: Number(form.get("sueldo_base") || 0),
    estado: String(form.get("estado") || "activo")
  };

  if (!payload.nombre) {
    toast("Debes ingresar el nombre del trabajador.");
    return;
  }

  const { data, error } = await supabase
    .from("trabajadores")
    .insert(payload)
    .select("*")
    .single();

  if (error) {
    console.error(error);
    toast("No se pudo guardar el trabajador.");
    return;
  }

  setTrabajadores(prev => [data, ...prev]);
  setModalTrabajador(false);
};

const subirDocumentoTrabajador = async (e) => {
  e.preventDefault();
  if (!trabajadorArchivo?.trabajador || !trabajadorArchivo?.file) return;

  const form = new FormData(e.currentTarget);
  const tipo = String(form.get("tipo") || "documento");
  const trabajador = trabajadorArchivo.trabajador;
  const file = trabajadorArchivo.file;
  const cleanName = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
  const path = `${trabajador.id}/${Date.now()}_${cleanName}`;

  const { error: uploadError } = await supabase.storage
    .from("rrhh")
    .upload(path, file, { upsert:false });

  if (uploadError) {
    console.error(uploadError);
    toast("No se pudo subir el archivo. Revisa que exista el bucket rrhh en Supabase Storage.");
    return;
  }

  const { data: publicData } = supabase.storage.from("rrhh").getPublicUrl(path);

  const { data, error } = await supabase
    .from("trabajador_documentos")
    .insert({ trabajador_id: trabajador.id, tipo, nombre_archivo: file.name, storage_path: path, url_publica: publicData?.publicUrl || "" })
    .select("*")
    .single();

  if (error) {
    console.error(error);
    toast("El archivo subió, pero no se pudo registrar en la ficha del trabajador.");
    return;
  }

  setDocumentosTrabajadores(prev => [data, ...prev]);
  setTrabajadorArchivo(null);
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
  const estadoBase = estadoProduccionDesdeNota(n);

  if (estadoBase.clave !== "terminada" && estadoBase.clave !== "entregada" && estaAtrasada(n.fecha_entrega_estimada)) {
    return {
      texto: "🔴 Atrasada",
      color: COLORS.danger
    };
  }

  return estadoBase;
};
const coincideFiltroProduccion = (n, filtro) => {
  const estado = estadoProduccion(n).texto;

  if (filtro === "todos") return true;
  if (filtro === "atrasadas") return estado.includes("Atrasada");
  if (filtro === "hoy") return venceHoy(n.fecha_entrega_estimada) && !n.media_cana;
  if (filtro === "manana") return venceManana(n.fecha_entrega_estimada) && !n.media_cana;
  if (filtro === "proceso") return estado.includes("En proceso");
  if (filtro === "sin_iniciar") return estado.includes("Sin iniciar");
  if (filtro === "listas") return estado.includes("Lista");

  return true;
};

const prioridadProduccion = (n) => {
  if (estaAtrasada(n.fecha_entrega_estimada) && !n.media_cana) return 1;
  if (venceHoy(n.fecha_entrega_estimada) && !n.media_cana) return 2;
  if (venceManana(n.fecha_entrega_estimada) && !n.media_cana) return 3;
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

const numeroDocumentoOrden = (valor) => {
  const match = String(valor || "").match(/\d+/g);
  if (!match) return 0;
  return Number(match.join("")) || 0;
};

const ordenesCompraImportadas = documentosImportados
  .filter(d => {
    const tipo = String(d.tipo || "").toUpperCase();
    const origen = String(d.origen || "").toLowerCase();
    return tipo === "OC" || origen === "oc_excel";
  })
  .sort((a, b) => {
    const numeroA = numeroDocumentoOrden(a.documento || a.folio || a.clave);
    const numeroB = numeroDocumentoOrden(b.documento || b.folio || b.clave);
    if (numeroB !== numeroA) return numeroB - numeroA;
    return new Date(b.created_at || b.fecha || 0) - new Date(a.created_at || a.fecha || 0);
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
    toast("Error al guardar el producto.");
    return;
  }

  setProductosInventario(prev => [...prev, data[0]].sort((a,b) => a.nombre.localeCompare(b.nombre)));
  setModalNuevoProducto(false);
};
const guardarMovimientoInventario = async ({ producto, tipo, cantidad }) => {
  const cantidadNumero = Number(cantidad);

  if (!cantidadNumero || cantidadNumero <= 0) {
    toast("Debes ingresar una cantidad válida.");
    return;
  }

  const stockAnterior = Number(producto.stock_actual || 0);

  if (tipo === "salida" && cantidadNumero > stockAnterior) {
    toast("No puedes descontar más stock del disponible.");
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
    toast("Error al actualizar el stock.");
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
    toast("El stock cambió, pero hubo un error guardando el historial.");
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
    toast("Debes ingresar al menos un sobrante válido.");
    return;
  }

  const { data, error } = await supabase
    .from("inventario_laminado_sobrantes")
    .insert(sobrantesValidos)
    .select();

  if (error) {
    console.error(error);
    toast("Error al guardar los sobrantes.");
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
    toast("Error al marcar el sobrante como usado.");
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
    toast("Error al editar el producto.");
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
    toast("No se generó el asiento automático porque Debe y Haber no cuadran.");
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
    toast("La operación se guardó, pero no se pudo crear el asiento contable automático.");
    return false;
  }

  setAsientosContables(prev => [...(data || []), ...prev]);
  return true;
};

const crearAsientoCompraAutomatica = async ({ fecha, documento, proveedor, totalCompra, estadoPago, detalleBase, cuentaCompra = "Compras" }) => {
  const { total, neto, iva } = calcularNetoIvaDesdeTotal(totalCompra);
  if (!total) return true;

  const doc = documento || `COMPRA-${Date.now()}`;
  const detalle = detalleBase || `Compra ${proveedor || "proveedor"}`;
  const cuentaHaber = estadoPago === "pendiente" ? "Proveedores" : (estadoPago === "caja" ? "Caja" : "Banco");

  return guardarAsientosAutomaticos([
    { fecha, detalle, desglose:"Compra automática", documento:doc, definicion:cuentaCompra || "Compras", debe:neto, haber:0 },
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

const crearAsientoAbonoNVAutomatico = async ({ fecha, numero, cliente, monto, medioPago = "transferencia", tipoDocumento = "nv" }) => {
  const total = Math.round(Number(monto || 0));
  if (!total) return true;

  const medio = String(medioPago || "transferencia").toLowerCase();
  const cuentaDebe = medio === "efectivo" ? "Caja" : medio === "cheque" ? "Cheques por cobrar" : "Banco";
  const etiquetaDocumento = String(tipoDocumento || "nv").toLowerCase() === "barran" ? "Barrán" : "NV";
  const prefijoDocumento = String(tipoDocumento || "nv").toLowerCase() === "barran" ? "ABONO-BARRAN" : "ABONO-NV";
  const documento = `${prefijoDocumento}-${numero}-${Date.now()}`;
  const detalle = `Abono ${etiquetaDocumento} ${numero} - ${cliente || "cliente"} (${cuentaDebe})`;

  return guardarAsientosAutomaticos([
    { fecha, detalle, desglose:"Abono cliente", documento, definicion:cuentaDebe, debe:total, haber:0 },
    { fecha, detalle, desglose:"Abono cliente", documento, definicion:"Clientes", debe:0, haber:total }
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


const normalizarClaveDocumento = (valor) => String(valor || "")
  .trim()
  .toUpperCase()
  .replace(/\s+/g, "")
  .replace(/[^A-Z0-9K\-]/g, "");

const existeDocumentoImportado = async (clave) => {
  if (!clave) return false;

  const { data, error } = await supabase
    .from("documentos_importados")
    .select("id")
    .eq("clave", clave)
    .maybeSingle();

  if (error) {
    console.error(error);
    toast("No se pudo verificar si el documento ya fue importado. Ejecuta primero el SQL de documentos_importados en Supabase.");
    return true;
  }

  return !!data;
};


const existeOCImportadaLocal = (documento) => {
  const docNormalizado = normalizarClaveDocumento(documento);
  if (!docNormalizado) return false;

  return documentosImportados.some(d => {
    const tipo = String(d.tipo || "").toUpperCase();
    const origen = String(d.origen || "").toLowerCase();
    const esOC = tipo === "OC" || origen === "oc_excel";
    if (!esOC) return false;

    const docGuardado = normalizarClaveDocumento(d.documento || d.folio || "");
    const claveGuardada = normalizarClaveDocumento(d.clave || "");

    return docGuardado === docNormalizado || claveGuardada === `OC-${docNormalizado}` || claveGuardada.endsWith(`-${docNormalizado}`);
  });
};

const registrarDocumentoImportado = async ({ clave, tipo, proveedor, documento, fecha, total, origen }) => {
  if (!clave) return true;

  const { error } = await supabase
    .from("documentos_importados")
    .insert([{
      clave,
      tipo: tipo || "documento",
      proveedor: proveedor || "",
      documento: documento || "",
      fecha: fecha || new Date().toISOString().split("T")[0],
      total: Number(total || 0),
      origen: origen || "app"
    }]);

  if (error) {
    console.error(error);
    toast("La operación se guardó, pero no se pudo registrar el documento como importado. Revisa la tabla documentos_importados.");
    return false;
  }

  const nuevoDocumento = {
    clave,
    tipo: tipo || "documento",
    proveedor: proveedor || "",
    documento: documento || "",
    fecha: fecha || new Date().toISOString().split("T")[0],
    total: Number(total || 0),
    origen: origen || "app",
    created_at: new Date().toISOString()
  };

  setDocumentosImportados(prev => [nuevoDocumento, ...prev]);

  return true;
};

const subirProductosAInventarioDesdeDocumento = async (productos, origen = "documento_compra") => {
  for (const item of productos) {
    const nombre = String(item.nombre || "").trim().toUpperCase();
    const cantidad = Number(item.cantidad || 0);

    if (!nombre || cantidad <= 0) continue;

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
        toast(`Error actualizando ${nombre}`);
        return false;
      }

      await supabase
        .from("inventario_movimientos")
        .insert([{
          producto_id: productoExistente.id,
          tipo: "entrada",
          cantidad,
          stock_anterior: stockAnterior,
          stock_nuevo: stockNuevo,
          origen
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
        toast(`Error creando ${nombre}`);
        return false;
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
          origen
        }]);

      setProductosInventario(prev =>
        [...prev, nuevoProducto].sort((a,b) => a.nombre.localeCompare(b.nombre))
      );
    }
  }

  return true;
};

const textoXml = (xmlDoc, tag) => {
  const el = xmlDoc.getElementsByTagName(tag)?.[0];
  return String(el?.textContent || "").trim();
};

const leerFacturaXml = async (file) => {
  const texto = await file.text();
  const xmlDoc = new DOMParser().parseFromString(texto, "text/xml");

  if (xmlDoc.getElementsByTagName("parsererror").length) {
    throw new Error("El XML no se pudo leer correctamente.");
  }

  const tipoDte = textoXml(xmlDoc, "TipoDTE");
  const folio = textoXml(xmlDoc, "Folio");
  const fecha = textoXml(xmlDoc, "FchEmis") || new Date().toISOString().split("T")[0];
  const rutProveedor = textoXml(xmlDoc, "RUTEmisor");
  const proveedor = textoXml(xmlDoc, "RznSoc") || "Proveedor";
  const neto = Number(textoXml(xmlDoc, "MntNeto") || 0);
  const iva = Number(textoXml(xmlDoc, "IVA") || 0);
  const total = Number(textoXml(xmlDoc, "MntTotal") || 0);

  if (!tipoDte || !folio || !rutProveedor || !total) {
    throw new Error("No se pudo leer tipo DTE, folio, proveedor o total del XML.");
  }

  const detalles = Array.from(xmlDoc.getElementsByTagName("Detalle"))
    .map((d) => ({
      nombre: String(d.getElementsByTagName("NmbItem")?.[0]?.textContent || "").trim().toUpperCase(),
      cantidad: Number(d.getElementsByTagName("QtyItem")?.[0]?.textContent || 0),
      precio: Number(d.getElementsByTagName("PrcItem")?.[0]?.textContent || 0),
      total: Number(d.getElementsByTagName("MontoItem")?.[0]?.textContent || 0)
    }))
    .filter(d => d.nombre && d.cantidad > 0);

  if (detalles.length === 0) {
    throw new Error("El XML no trae productos/detalle legible.");
  }

  return {
    tipoDte,
    folio,
    fecha,
    rutProveedor,
    proveedor,
    neto,
    iva,
    total,
    detalles,
    clave: `${tipoDte}-${normalizarClaveDocumento(rutProveedor)}-${normalizarClaveDocumento(folio)}`,
    documento: `DTE-${tipoDte}-${folio}`,
    archivo: file.name
  };
};

const importarFacturaXml = async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const factura = await leerFacturaXml(file);
    const yaExiste = await existeDocumentoImportado(factura.clave);

    if (yaExiste) {
      toast(`Esta factura XML ya fue importada anteriormente.\n\nDocumento: ${factura.documento}\nProveedor: ${factura.proveedor}`);
      event.target.value = "";
      return;
    }

    setPreviewFacturaXml(factura);
    setModalFacturaXml(true);
  } catch (error) {
    console.error(error);
    toast(error.message || "No se pudo leer la factura XML.");
  }

  event.target.value = "";
};

const confirmarImportacionFacturaXml = async ({ modo, estadoPago, cuentaCompra }) => {
  if (!previewFacturaXml) return;

  const factura = previewFacturaXml;
  const yaExiste = await existeDocumentoImportado(factura.clave);

  if (yaExiste) {
    toast(`Esta factura ya fue importada anteriormente.\n\n${factura.documento} · ${factura.proveedor}`);
    setModalFacturaXml(false);
    setPreviewFacturaXml(null);
    return;
  }

  if (modo === "inventario_contabilidad") {
    const okInventario = await subirProductosAInventarioDesdeDocumento(
      factura.detalles,
      "factura_xml"
    );

    if (!okInventario) return;
  }

  await crearAsientoCompraAutomatica({
    fecha: factura.fecha,
    documento: factura.documento,
    proveedor: factura.proveedor,
    totalCompra: factura.total,
    estadoPago: estadoPago || "pagado",
    cuentaCompra: cuentaCompra || (modo === "inventario_contabilidad" ? "Compras" : "Gastos generales"),
    detalleBase: `Factura XML ${factura.folio} ${factura.proveedor}`.trim()
  });

  await registrarDocumentoImportado({
    clave: factura.clave,
    tipo: `DTE-${factura.tipoDte}`,
    proveedor: factura.proveedor,
    documento: factura.documento,
    fecha: factura.fecha,
    total: factura.total,
    origen: modo === "inventario_contabilidad" ? "xml_inventario_contabilidad" : "xml_contabilidad"
  });

  setPreviewFacturaXml(null);
  setModalFacturaXml(false);
  toast("Factura XML importada correctamente.");
};

const leerProductosDesdeOC = async (file) => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const hojaNombre = workbook.SheetNames[0];
  const hoja = workbook.Sheets[hojaNombre];

  const filas = XLSX.utils.sheet_to_json(hoja, {
    header: 1,
    defval: ""
  });

  const documentoDetectado = String(hoja?.["F13"]?.v || "").trim() || String(file.name || "").split(" ")[0];

  if (!documentoDetectado) {
    throw new Error(`No se pudo leer el número de OC en F13 de ${file.name}`);
  }

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
    throw new Error(`No se encontró la tabla de productos en ${file.name}`);
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
      cantidad,
      documento_origen: documentoDetectado,
      archivo_origen: file.name
    });
  }

  if (productosEncontrados.length === 0) {
    throw new Error(`No se encontraron productos válidos en ${file.name}`);
  }

  const proveedorDetectado = String(file.name || "")
    .replace(/\.[^.]+$/, "")
    .replace(documentoDetectado, "")
    .trim();

  return {
    documento: documentoDetectado,
    proveedor: proveedorDetectado,
    archivo: file.name,
    productos: productosEncontrados,
    clave: `OC-${normalizarClaveDocumento(documentoDetectado)}`
  };
};

const importarOrdenCompraExcel = async (e) => {
  const files = Array.from(e.target.files || []);
  if (files.length === 0) return;

  const productosTodos = [];
  const documentosValidos = [];
  const errores = [];
  const duplicados = [];

  for (const file of files) {
    try {
      const oc = await leerProductosDesdeOC(file);

      if (existeOCImportadaLocal(oc.documento) || await existeDocumentoImportado(oc.clave)) {
        duplicados.push(`${oc.documento} (${file.name})`);
        continue;
      }

      productosTodos.push(...oc.productos);
      documentosValidos.push({
        documento: oc.documento,
        proveedor: oc.proveedor,
        archivo: oc.archivo,
        clave: oc.clave
      });
    } catch (error) {
      console.error(error);
      errores.push(`${file.name}: ${error.message || "Error desconocido"}`);
    }
  }

  if (productosTodos.length === 0) {
    toast(
      `No se pudo importar ninguna orden de compra.` +
      (duplicados.length ? `\n\nDuplicadas:\n${duplicados.slice(0, 8).join("\n")}` : "") +
      (errores.length ? `\n\nErrores:\n${errores.slice(0, 8).join("\n")}` : "")
    );
    e.target.value = "";
    return;
  }

  const agrupados = productosTodos.reduce((acc, item) => {
    const existente = acc.find(p => p.nombre === item.nombre);

    if (existente) {
      existente.cantidad += item.cantidad;
    } else {
      acc.push({ ...item });
    }

    return acc;
  }, []);

  if (errores.length || duplicados.length) {
    toast(
      `Se leyeron ${documentosValidos.length} OC nueva(s).` +
      (duplicados.length ? `\n\nDuplicadas omitidas:\n${duplicados.slice(0, 8).join("\n")}` : "") +
      (errores.length ? `\n\nCon error:\n${errores.slice(0, 8).join("\n")}` : "")
    );
  }

  setPreviewOC(agrupados);
  setPreviewOCDocumentos(documentosValidos);
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

  if (previewOCDocumentos.length === 0) {
    toast("No hay OC detectadas para registrar. Vuelve a importar los archivos.");
    return;
  }

  for (const doc of previewOCDocumentos) {
    if (existeOCImportadaLocal(doc.documento) || await existeDocumentoImportado(doc.clave)) {
      toast(`La OC ${doc.documento} ya fue importada anteriormente. Se canceló el ingreso para evitar duplicar stock.`);
      return;
    }
  }

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
        toast(`Error actualizando ${nombre}`);
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
        toast(`Error creando ${nombre}`);
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

  const documentoContable = previewOCDocumentos.length === 1
    ? previewOCDocumentos[0].documento
    : `OC-MASIVA-${new Date().toISOString().slice(0,10)}`;

  if (Number(datosCompra.totalCompra || 0) > 0) {
    await crearAsientoCompraAutomatica({
      fecha: datosCompra.fecha || new Date().toISOString().split("T")[0],
      documento: documentoContable,
      proveedor: datosCompra.proveedor || (previewOCDocumentos[0]?.proveedor || "Proveedor"),
      totalCompra: datosCompra.totalCompra,
      estadoPago: datosCompra.estadoPago || "pagado",
      detalleBase: `Compra ${documentoContable} ${datosCompra.proveedor || previewOCDocumentos[0]?.proveedor || ""}`.trim()
    });
  }

  for (const doc of previewOCDocumentos) {
    await registrarDocumentoImportado({
      clave: doc.clave,
      tipo: "OC",
      proveedor: datosCompra.proveedor || doc.proveedor || "",
      documento: doc.documento,
      fecha: datosCompra.fecha || new Date().toISOString().split("T")[0],
      total: previewOCDocumentos.length === 1 ? Number(datosCompra.totalCompra || 0) : 0,
      origen: "oc_excel"
    });
  }

  setPreviewOC([]);
  setPreviewOCDocumentos([]);
  setModalPreviewOC(false);
  toast("Orden(es) de compra importada(s) correctamente.");
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
    toast("No se pudo leer la primera hoja del Excel.");
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
    toast("No se pudo leer el número o total de la cotización de láminas.");
    event.target.value = "";
    return;
  }

  const duplicada = ventasLaminas.some(v => String(v.numero) === String(numero));
  if (duplicada) {
    toast("Esta venta de láminas ya existe.");
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
    toast("No se encontró el detalle de láminas en el Excel.");
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
    toast("Error al guardar la venta de láminas. Revisa si las tablas están creadas en Supabase.");
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
      toast("La venta se guardó, pero hubo un error guardando el detalle.");
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
  toast("Venta de láminas importada correctamente.");
};

const guardarCostoVentaLaminas = async (venta, costoCompraTotal) => {
  const { data, error } = await supabase
    .from("ventas_laminas")
    .update({ costo_compra_total: costoCompraTotal })
    .eq("id", venta.id)
    .select();

  if (error) {
    console.error(error);
    toast("Error al guardar el costo de compra.");
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
      toast("No se pudo actualizar el asiento contable.");
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
      toast("No se pudo guardar el asiento contable. Revisa que la tabla exista en Supabase.");
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
    toast("Los asientos no cuadran. Revisa los montos antes de guardar.");
    return;
  }

  const { data, error } = await supabase
    .from("asientos_contables")
    .insert(payload)
    .select();

  if (error) {
    console.error(error);
    toast("No se pudieron guardar los asientos automáticos.");
    return;
  }

  setAsientosContables(prev => [...data, ...prev]);
  setModalGestionContable(null);
  toast("Gestión contable registrada correctamente.");
};

const eliminarAsientoContable = async (asiento) => {
  const confirmar = await confirmDialog("¿Seguro que quieres eliminar este asiento contable?\n\nEsta acción no se puede deshacer.");
  if (!confirmar) return;

  const { error } = await supabase
    .from("asientos_contables")
    .delete()
    .eq("id", asiento.id);

  if (error) {
    console.error(error);
    toast("No se pudo eliminar el asiento contable.");
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

const calcularAbonadoDocumento = (nota) => {
  const idNota = Number(String(nota.id).replace("supabase-", ""));
  const abonadoDesdeNota = Number(nota.abono || nota.abonado || 0);
  const abonadoDesdeHistorial = abonosNV
    .filter(a => Number(a.nota_venta_id) === idNota)
    .reduce((sum, a) => sum + Number(a.monto || 0), 0);
  return Math.max(abonadoDesdeNota, abonadoDesdeHistorial);
};

const obtenerEstadoCxc = (nota, saldo) => {
  if (saldo <= 0) return { key:"pagada", label:"Pagada", icon:"🟢", color:COLORS.success };
  const entregada = String(nota.proceso || "").toLowerCase() === "entregado";
  if (entregada && clienteTieneCreditoAutorizado(nota.cliente)) {
    return { key:"credito_autorizado", label:"Crédito autorizado", icon:"🟣", color:"#a78bfa" };
  }
  if (entregada) {
    return { key:"entregada_por_confirmar", label:"Entregada por confirmar", icon:"🔴", color:COLORS.danger };
  }
  return { key:"saldo_pendiente", label:"Saldo pendiente", icon:"🟡", color:COLORS.warning };
};

const detalleCxcNotasTodas = notas
  .map(n => {
    const total = Number(n.total || 0);
    const abonado = calcularAbonadoDocumento(n);
    const saldo = Math.max(total - abonado, 0);
    const estado = obtenerEstadoCxc(n, saldo);
    const tipo = (n.tipo_documento || "nv") === "barran" ? "Barrán" : "NV";
    const detalles = detallesNotasVenta.filter(d => String(d.nota_venta_numero) === String(n.numero));

    return {
      nota: n,
      id: n.id,
      numero: n.numero || n.nota_venta || n.id,
      tipo,
      cliente: n.cliente || "Sin cliente",
      fecha: n.fecha,
      total,
      abonado,
      saldo,
      estado,
      proceso: n.proceso || "en espera",
      fechaEntrega: n.fecha_entrega_estimada || "",
      detalles,
      detalleTexto: detalles.slice(0, 2).map(d => `${d.cantidad || 0} x ${d.descripcion || d.material || "item"} ${d.color ? `· ${d.color}` : ""}`).join(" | ")
    };
  })
  .filter(n => n.saldo > 0);

const detalleCxcNotas = detalleCxcNotasTodas.filter(n =>
  String(n.fecha || "").slice(0, 7) === mesCxc
);

const totalCxcNotas = detalleCxcNotas.reduce((sum, n) => sum + Number(n.saldo || 0), 0);
const cxcSaldosPendientes = detalleCxcNotas.filter(n => n.estado.key === "saldo_pendiente");
const cxcEntregadasPorConfirmar = detalleCxcNotas.filter(n => n.estado.key === "entregada_por_confirmar");
const cxcCreditoAutorizado = detalleCxcNotas.filter(n => n.estado.key === "credito_autorizado");

const totalCxcSaldosPendientes = cxcSaldosPendientes.reduce((sum, n) => sum + Number(n.saldo || 0), 0);
const totalCxcEntregadasPorConfirmar = cxcEntregadasPorConfirmar.reduce((sum, n) => sum + Number(n.saldo || 0), 0);
const totalCxcCreditoAutorizado = cxcCreditoAutorizado.reduce((sum, n) => sum + Number(n.saldo || 0), 0);

const detalleCxcFiltrado = detalleCxcNotas.filter(item => {
  const q = String(busquedaCxcCartola || "").toLowerCase().trim();
  if (!q) return true;
  return [
    item.numero,
    item.tipo,
    item.cliente,
    item.fecha,
    item.estado.label,
    item.detalleTexto
  ].join(" ").toLowerCase().includes(q);
});

const categoriasCartola = [
  { nombre:"Abono CxC", tipo:"ingreso", descripcion:"Pago de cliente asociado a NV o Barrán" },
  { nombre:"Préstamo socio", tipo:"ingreso", descripcion:"Dinero aportado por Franco/familia para caja" },
  { nombre:"Venta sin documento", tipo:"ingreso", descripcion:"Ingreso no asociado a CxC" },
  { nombre:"Factura compra", tipo:"egreso", descripcion:"Compra facturada; se considera flujo, no duplica resumen SII" },
  { nombre:"Compra sin factura", tipo:"egreso", descripcion:"Compra o gasto sin documento tributario" },
  { nombre:"Remuneraciones", tipo:"egreso", descripcion:"Sueldos y pagos a trabajadores" },
  { nombre:"Imposiciones", tipo:"egreso", descripcion:"Previred, AFP, salud, cesantía" },
  { nombre:"Arriendo", tipo:"egreso", descripcion:"Arriendo local/bodega" },
  { nombre:"Bencina", tipo:"egreso", descripcion:"Combustible" },
  { nombre:"TAG", tipo:"egreso", descripcion:"Autopistas" },
  { nombre:"Servicios básicos", tipo:"egreso", descripcion:"Luz, agua, internet, telefonía" },
  { nombre:"Deuda antigua", tipo:"egreso", descripcion:"Pagos de deudas históricas" },
  { nombre:"Retiro socio", tipo:"egreso", descripcion:"Retiros personales o familiares" },
  { nombre:"Transferencia interna", tipo:"neutro", descripcion:"Movimiento entre cuentas propias" },
  { nombre:"No sumar", tipo:"neutro", descripcion:"Movimiento duplicado o que no corresponde analizar" },
  { nombre:"Otro", tipo:"otro", descripcion:"Clasificación manual general" }
];

const categoriasCartolaFiltradas = categoriasCartola.filter(c => {
  const q = busquedaCategoriaCartola.toLowerCase().trim();
  if (!q) return true;
  return `${c.nombre} ${c.descripcion}`.toLowerCase().includes(q);
});

const opcionesCxcCartola = detalleCxcNotas.filter(item => {
  const q = busquedaCxcCartola.toLowerCase().trim();
  if (!q) return true;
  return `${item.tipo} ${item.numero} ${item.cliente} ${item.saldo} ${item.detalleTexto}`.toLowerCase().includes(q);
});

const eliminarVentaLaminas = async (venta) => {
  const confirmar = await confirmDialog(
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
    toast("No se pudo eliminar la venta de láminas.");
    return;
  }

  setVentasLaminas(prev => prev.filter(v => v.id !== venta.id));
  setDetallesVentasLaminas(prev => prev.filter(d => d.venta_lamina_id !== venta.id));

  if (modalVentaLaminas?.id === venta.id) {
    setModalVentaLaminas(null);
  }

  toast("Venta de láminas eliminada correctamente.");
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


const numeroProduccion = (n) => {
  const texto = String(n?.numero || "");
  const match = texto.match(/\d+/);
  return match ? Number(match[0]) : 0;
};

const ordenarProduccion = (items) => {
  const lista = [...items];

  if (ordenProduccion === "fecha_entrega") {
    return lista.sort((a,b) => {
      const prioridadA = prioridadProduccion(a);
      const prioridadB = prioridadProduccion(b);

      if (prioridadA !== prioridadB) return prioridadA - prioridadB;

      if (!a.fecha_entrega_estimada && !b.fecha_entrega_estimada) {
        return numeroProduccion(b) - numeroProduccion(a);
      }
      if (!a.fecha_entrega_estimada) return 1;
      if (!b.fecha_entrega_estimada) return -1;

      const diferenciaFecha = new Date(a.fecha_entrega_estimada) - new Date(b.fecha_entrega_estimada);
      if (diferenciaFecha !== 0) return diferenciaFecha;

      return numeroProduccion(b) - numeroProduccion(a);
    });
  }

  return lista.sort((a,b) => {
    const diferenciaNumero = numeroProduccion(b) - numeroProduccion(a);
    if (diferenciaNumero !== 0) return diferenciaNumero;
    return String(b.numero || "").localeCompare(String(a.numero || ""), "es-CL", { numeric:true });
  });
};

const trabajosProduccionBase = notas
  .filter(n => n.proceso !== "entregado")
  .filter(n => coincideFiltroProduccion(n, filtroProduccion));

const trabajosProduccionNV = ordenarProduccion(
  trabajosProduccionBase.filter(n => (n.tipo_documento || "nv") !== "barran")
);

const trabajosProduccionBarranes = ordenarProduccion(
  trabajosProduccionBase.filter(n => (n.tipo_documento || "nv") === "barran")
);

const renderTarjetaProduccion = (n) => (
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
        <b style={{ color:COLORS.success }}>{(n.tipo_documento || "nv") === "barran" ? "Barrán #" : "NV#"}{n.numero}</b>
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
      {n.media_cana && <span>✅ Media caña</span>}
    </div>

    {n.produccion_observaciones && (
      <p style={{ marginTop:10, color:COLORS.warning }}>
        Obs: {n.produccion_observaciones}
      </p>
    )}

    <div style={{ marginTop:12, display:"flex", justifyContent:"flex-end" }}>
      <button
        onClick={(e) => {
          e.stopPropagation();
          marcarComoEntregadaProduccion(n);
        }}
        style={{
          background: COLORS.success,
          color: "#fff",
          border: "none",
          borderRadius: 8,
          padding: "7px 11px",
          fontWeight: 700,
          cursor: "pointer",
          fontSize: 12
        }}
      >
        Entregada
      </button>
    </div>
  </div>
);


  const cambiarMesCalendario = (delta) => {
    const [year, month] = mesCalendario.split("-").map(Number);
    const fecha = new Date(year, month - 1 + delta, 1);
    setMesCalendario(`${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, "0")}`);
  };

  const obtenerAbonosCuenta = (cuentaId) => abonosCuentasPagar.filter(a => String(a.cuenta_pagar_id) === String(cuentaId));
  const saldoCuentaPagar = (cuenta) => Math.max(Number(cuenta.monto_total || 0) - obtenerAbonosCuenta(cuenta.id).reduce((sum, a) => sum + Number(a.monto || 0), 0), 0);

  const eventosCalendario = [
    ...notas
      .filter(n => n.fecha_entrega_estimada && String(n.proceso || "").toLowerCase() !== "entregado")
      .map(n => ({
        fecha: n.fecha_entrega_estimada,
        tipo: (n.tipo_documento || "nv") === "barran" ? "Barrán" : "NV",
        titulo: `${(n.tipo_documento || "nv") === "barran" ? "Barrán" : "NV"} ${n.numero}`,
        subtitulo: n.cliente,
        color: (n.tipo_documento || "nv") === "barran" ? COLORS.success : "#60a5fa",
        onClick: () => setModalProduccion(n)
      })),
    ...cuentasPagar
      .filter(c => c.fecha_vencimiento && saldoCuentaPagar(c) > 0)
      .map(c => ({
        fecha: c.fecha_vencimiento,
        tipo: "Pago",
        titulo: `${c.tipo || "Cuenta"}: ${c.nombre}`,
        subtitulo: `${fmt(saldoCuentaPagar(c))} pendiente`,
        color: COLORS.warning,
        onClick: () => setModalAbonoCuentaPagar(c)
      }))
  ];

  const [calYear, calMonth] = mesCalendario.split("-").map(Number);
  const diasMesCalendario = new Date(calYear, calMonth, 0).getDate();
  const primerDiaCalendario = new Date(calYear, calMonth - 1, 1).getDay();
  const celdasVaciasInicio = primerDiaCalendario === 0 ? 6 : primerDiaCalendario - 1;
  const eventosDelMes = eventosCalendario.filter(e => String(e.fecha || "").startsWith(mesCalendario));
  const pagosPendientesMes = cuentasPagar.filter(c => String(c.fecha_vencimiento || "").startsWith(mesCalendario) && saldoCuentaPagar(c) > 0);
  const totalPagosPendientesMes = pagosPendientesMes.reduce((sum, c) => sum + saldoCuentaPagar(c), 0);
  const entregasPendientesMes = notas.filter(n => String(n.fecha_entrega_estimada || "").startsWith(mesCalendario) && String(n.proceso || "").toLowerCase() !== "entregado");
  const cartolaFiltrada = cartolaMovimientos.filter(m => {
    const coincideMes = !mesCartola || String(m.fecha || "").startsWith(mesCartola);
    const texto = `${m.descripcion || ""} ${m.categoria || ""} ${m.estado || ""}`.toLowerCase();
    const coincideBusqueda = !busquedaCartola || texto.includes(busquedaCartola.toLowerCase());
    return coincideMes && coincideBusqueda;
  });
  const cartolaPorRevisar = cartolaFiltrada.filter(m => (m.estado || "por_revisar") !== "conciliado");
  const totalIngresosCartola = cartolaFiltrada.reduce((sum, m) => sum + Number(m.abono || 0), 0);
  const totalEgresosCartola = cartolaFiltrada.reduce((sum, m) => sum + Number(m.cargo || 0), 0);

  const resumenFinancieroCartola = cartolaFiltrada.reduce((acc, m) => {
    const categoria = String(m.categoria || "Por revisar");
    const cargo = Number(m.cargo || 0);
    const abono = Number(m.abono || 0);

    if (!acc[categoria]) acc[categoria] = { categoria, ingresos:0, egresos:0, movimientos:0 };
    acc[categoria].ingresos += abono;
    acc[categoria].egresos += cargo;
    acc[categoria].movimientos += 1;

    return acc;
  }, {});

  const resumenCartolaCategorias = Object.values(resumenFinancieroCartola)
    .sort((a, b) => (b.ingresos + b.egresos) - (a.ingresos + a.egresos));

  const tabsPublicos=[
    {key:"produccion",label:`🏭 Producción`},
    {key:"inventario",label:`📦 Inventario`},
    {key:"calendario",label:`📅 Calendario (${eventosDelMes.length})`},
    {key:"quotes",label:`📋 Cotizaciones (${filteredQuotes.length})`},
    {key:"sales",label:`✅ Notas de Venta (${filteredNotas.length})`},
    {key:"barranes",label:`🧾 Barranes (${filteredBarranes.length})`},
    {key:"venta_laminas",label:`🧾 Venta de Láminas (${ventasLaminasDelMes.length})`},
    {key:"cuentas_pagar",label:`💸 Cuentas por pagar (${pagosPendientesMes.length})`},
    {key:"dashboard",label:"📊 Resumen"},
  ];

  const tabsAdmin=[
    {key:"control_calidad",label:`🔍 Control de Calidad (${erroresProduccion.length})`},
    {key:"cartola",label:`🏦 Cartola (${cartolaPorRevisar.length})`},
    {key:"cxc",label:`💰 CxC (${detalleCxcNotas.length})`},
    {key:"contabilidad",label:`📚 Contabilidad (${asientosContablesFiltrados.length})`},
    {key:"rrhh",label:`👷 RRHH (${trabajadores.length})`},
  ];

  // Filtrar tabs según permisos del usuario actual
  const seccionesUsuario = usuarioActual?.secciones || [];
  const tieneAccesoTotal = seccionesUsuario.includes("all");
  const todosLosTabs = [...tabsPublicos, ...tabsAdmin];
  const tabs = tieneAccesoTotal
    ? todosLosTabs
    : todosLosTabs.filter(t => seccionesUsuario.includes(t.key));

  return (
    <div style={{ minHeight:"100vh", background:COLORS.bg, color:COLORS.text, fontFamily:"'Trebuchet MS',sans-serif", paddingBottom:60 }}>
      <ToastContainer />
      <ConfirmContainer />

      {/* PANTALLA DE LOGIN OBLIGATORIO */}
      {!usuarioActual && (
        <div style={{ position:"fixed", inset:0, background:COLORS.bg, display:"flex", alignItems:"center", justifyContent:"center", zIndex:100002, padding:20 }}>
          <div style={{ width:"min(380px, 92vw)", background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:16, padding:32, boxSizing:"border-box" }}>
            <div style={{ textAlign:"center", marginBottom:24 }}>
              <div style={{ fontSize:34, marginBottom:8 }}>🪑</div>
              <h1 style={{ margin:0, fontSize:20, color:COLORS.accent }}>Muebles Santa Fe</h1>
              <p style={{ margin:"6px 0 0", fontSize:13, color:COLORS.muted }}>Panel de Conversión</p>
            </div>

            <label style={{ display:"block", marginBottom:6, fontSize:12, color:COLORS.muted }}>Usuario</label>
            <input
              type="text"
              value={loginUser}
              onChange={(e) => setLoginUser(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") document.getElementById("login-pass-input")?.focus(); }}
              placeholder="Tu nombre de usuario"
              autoFocus
              style={{ width:"100%", padding:"11px 13px", borderRadius:9, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, boxSizing:"border-box", marginBottom:14, fontSize:14 }}
            />

            <label style={{ display:"block", marginBottom:6, fontSize:12, color:COLORS.muted }}>Contraseña</label>
            <input
              id="login-pass-input"
              type="password"
              value={loginPass}
              onChange={(e) => setLoginPass(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && !cargandoLogin) hacerLoginAdmin(); }}
              placeholder="Tu contraseña"
              style={{ width:"100%", padding:"11px 13px", borderRadius:9, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, boxSizing:"border-box", marginBottom:20, fontSize:14 }}
            />

            <button
              onClick={hacerLoginAdmin}
              disabled={cargandoLogin}
              style={{ width:"100%", padding:"12px", background:cargandoLogin?COLORS.subtle:COLORS.accent, border:"none", color:"#0f0e0c", borderRadius:9, fontWeight:700, fontSize:15, cursor:cargandoLogin?"default":"pointer" }}
            >
              {cargandoLogin ? "Verificando…" : "Ingresar"}
            </button>
          </div>
        </div>
      )}

      {usuarioActual && (
      <>

      {cargando && (
        <div style={{ position:"fixed", inset:0, background:COLORS.bg, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:18, zIndex:100001 }}>
          <div style={{ width:46, height:46, border:`4px solid ${COLORS.subtle}`, borderTopColor:COLORS.accent, borderRadius:"50%", animation:"sf-spin 0.9s linear infinite" }} />
          <div style={{ color:COLORS.muted, fontSize:14, letterSpacing:0.5 }}>Cargando datos…</div>
          <style>{`@keyframes sf-spin { to { transform: rotate(360deg); } }`}</style>
        </div>
      )}

      <style>{`
        @media (max-width: 768px) {
          .solo-pc {
            display: none !important;
          }
        }
      `}</style>
 <div className="solo-pc">
  <div
  style={{
    margin: "16px auto",
    padding: 16,
    background: COLORS.card,
    border: `1px solid ${COLORS.border}`,
    borderRadius: 12,
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: 12,
    maxWidth: 720,
    width: "calc(100% - 32px)"
  }}
>
  <label style={{ padding:"14px 16px", background:COLORS.success, borderRadius:10, fontWeight:700, cursor:"pointer", color:"#fff", textAlign:"center" }}>
    Agregar NV
    <input type="file" accept=".xlsx,.xls" multiple onChange={importarExcel} style={{ display:"none" }} />
  </label>

  <label style={{ padding:"14px 16px", background:"#8b5cf6", borderRadius:10, fontWeight:700, cursor:"pointer", color:"#fff", textAlign:"center" }}>
    Agregar Barranes
    <input type="file" accept=".xlsx,.xls" multiple onChange={importarBarranesExcel} style={{ display:"none" }} />
  </label>

  <label style={{ padding:"14px 16px", background:COLORS.accent, borderRadius:10, fontWeight:700, cursor:"pointer", color:"#fff", textAlign:"center" }}>
    Agregar Cotización
    <input type="file" accept=".xlsx,.xls" multiple onChange={importarCotizacion} style={{ display:"none" }} />
  </label>

  <label style={{ padding:"14px 16px", background:COLORS.warning || "#d6b45f", borderRadius:10, fontWeight:700, cursor:"pointer", color:"#fff", textAlign:"center" }}>
    Agregar Venta Láminas
    <input type="file" accept=".xlsx,.xls" onChange={importarVentaLaminasExcel} style={{ display:"none" }} />
  </label>

  <label style={{ display:"none" }}>
    Importar cartola bancaria
    <input id="importar-cartola-bancaria" type="file" accept=".xlsx,.xls,.csv" multiple onChange={importarCartolaBancaria} style={{ display:"none" }} />
  </label>
</div>
</div>

{/* Barra de sesión de usuario */}
<div style={{ display:"flex", justifyContent:"flex-end", alignItems:"center", padding:"8px 16px", borderBottom:`1px solid ${COLORS.border}`, background:COLORS.surface, gap:12 }}>
  <span style={{ fontSize:12, color:COLORS.text }}>
    👤 <b style={{ color:COLORS.accent }}>{usuarioActual?.nombre || usuarioActual?.username}</b>
    {tieneAccesoTotal && <span style={{ marginLeft:6, fontSize:10, color:COLORS.success, fontWeight:700 }}>ADMIN</span>}
  </span>
  <button
    onClick={hacerLogoutAdmin}
    style={{
      background:COLORS.danger,
      border:"none",
      color:"#fff",
      borderRadius:6,
      padding:"6px 12px",
      fontWeight:700,
      cursor:"pointer",
      fontSize:12,
    }}
  >
    Cerrar sesión
  </button>
</div>
      

      <div style={{ display:"flex", gap:2, padding:"12px 12px 0", borderBottom:`1px solid ${COLORS.border}`, overflowX:"auto", WebkitOverflowScrolling:"touch" }}>
        {tabs.map(t=>(
          <button key={t.key} onClick={()=>setTab(t.key)} style={{ background:tab===t.key?COLORS.card:"transparent", border:`1px solid ${tab===t.key?COLORS.border:"transparent"}`, borderBottom:tab===t.key?`2px solid ${COLORS.accent}`:"2px solid transparent", borderRadius:"8px 8px 0 0", padding:"8px 16px", color:tab===t.key?COLORS.accent:COLORS.muted, cursor:"pointer", fontSize:15, fontWeight:tab===t.key?700:400, whiteSpace:"nowrap" }}>{t.label}</button>
        ))}
      </div>

      <div style={{ padding:"clamp(12px, 4vw, 22px)", maxWidth:1100, margin:"0 auto" }}>
        {tab==="control_calidad" && esAdmin && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20, gap:12 }}>
              <h2 style={{ margin:0, color:COLORS.accent }}>🔍 Control de Calidad</h2>
              <button
                onClick={() => setModalNuevoError(true)}
                style={{
                  background:COLORS.success,
                  border:"none",
                  color:"#fff",
                  borderRadius:8,
                  padding:"9px 16px",
                  fontWeight:700,
                  cursor:"pointer",
                  fontSize:13,
                }}
              >
                + Registrar Error
              </button>
            </div>

            {erroresProduccion.length === 0 ? (
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:20, textAlign:"center", color:COLORS.muted }}>
                <div style={{ fontSize:14 }}>Sin errores registrados.</div>
              </div>
            ) : (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                  <thead>
                    <tr style={{ background:COLORS.bg }}>
                      <th style={{ textAlign:"left", padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.muted, fontWeight:700 }}>NV</th>
                      <th style={{ textAlign:"left", padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.muted, fontWeight:700 }}>Cliente</th>
                      <th style={{ textAlign:"left", padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.muted, fontWeight:700 }}>Tipo de Error</th>
                      <th style={{ textAlign:"center", padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.muted, fontWeight:700 }}>Días Retrabaljo</th>
                      <th style={{ textAlign:"right", padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.muted, fontWeight:700 }}>Monto Pérdida</th>
                      <th style={{ textAlign:"center", padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.muted, fontWeight:700 }}>Acciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {erroresProduccion.map((err) => (
                      <tr key={err.id} style={{ background:COLORS.card, borderBottom:`1px solid ${COLORS.border}` }}>
                        <td style={{ padding:10, color:COLORS.accent, fontWeight:700 }}>{err.nota_venta_numero}</td>
                        <td style={{ padding:10, color:COLORS.text }}>{err.cliente}</td>
                        <td style={{ padding:10, color:COLORS.text }}>{err.tipo_error}</td>
                        <td style={{ padding:10, textAlign:"center", color:COLORS.warning, fontWeight:700 }}>{err.dias_retrabaljo}</td>
                        <td style={{ padding:10, textAlign:"right", color:COLORS.danger, fontWeight:700 }}>{fmt(err.monto_material_perdido)}</td>
                        <td style={{ padding:10, textAlign:"center" }}>
                          <button
                            onClick={() => {
                              setErrorEditando(err);
                              setFormEditError({ dias: err.dias_retrabaljo || 0, monto: err.monto_material_perdido || 0 });
                            }}
                            style={{ background:"none", border:"none", color:COLORS.accent, cursor:"pointer", fontSize:14 }}
                          >
                            ✏️
                          </button>
                          <button
                            onClick={() => eliminarError(err.id)}
                            style={{ background:"none", border:"none", color:COLORS.danger, cursor:"pointer", fontSize:14, marginLeft:8 }}
                          >
                            🗑️
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <div style={{ marginTop:20, padding:16, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10 }}>
              <div style={{ fontSize:12, color:COLORS.muted, marginBottom:8 }}>📊 Resumen del Mes</div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:12 }}>
                <div>
                  <div style={{ fontSize:11, color:COLORS.muted }}>Errores Registrados</div>
                  <div style={{ fontSize:18, fontWeight:700, color:COLORS.accent }}>{erroresProduccion.length}</div>
                </div>
                <div>
                  <div style={{ fontSize:11, color:COLORS.muted }}>Días de Retrabaljo</div>
                  <div style={{ fontSize:18, fontWeight:700, color:COLORS.warning }}>{erroresProduccion.reduce((s,e)=>s+(Number(e.dias_retrabaljo)||0),0)}</div>
                </div>
                <div>
                  <div style={{ fontSize:11, color:COLORS.muted }}>Pérdida Total</div>
                  <div style={{ fontSize:18, fontWeight:700, color:COLORS.danger }}>{fmt(erroresProduccion.reduce((s,e)=>s+(Number(e.monto_material_perdido)||0),0))}</div>
                </div>
              </div>
            </div>
          </div>
        )}
        {tab==="dashboard" && (
          <div>
            <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:16, marginBottom:18 }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, flexWrap:"wrap" }}>
                <div>
                  <h2 style={{ margin:"0 0 4px", color:COLORS.accent }}>📊 Resumen Comercial</h2>
                  <div style={{ fontSize:12, color:COLORS.muted }}>Periodo seleccionado: {rangoResumen.label}</div>
                </div>
                <select
                  value={filtroResumen}
                  onChange={(e) => setFiltroResumen(e.target.value)}
                  style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 12px", fontSize:13, minWidth:180 }}
                >
                  <option value="mes_actual">Mes actual</option>
                  <option value="mes_anterior">Mes anterior</option>
                  <option value="ultimos_3_meses">Últimos 3 meses</option>
                  <option value="ultimos_6_meses">Últimos 6 meses</option>
                  <option value="año_actual">Año actual</option>
                  <option value="todo">Todo</option>
                </select>
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:12, marginBottom:20 }}>
              <StatCard label="Conversión" value={`${resumenActual.conversion}%`} sub={`${resumenActual.vendidasPeriodo.length} de ${resumenActual.cotizacionesPeriodo.length} cotizaciones`} icon="🎯" color={COLORS.accent}/>
              <StatCard label="Conversión monto" value={`${resumenActual.conversionMonto}%`} sub={`${fmt(resumenActual.montoVendido)} vendido de ${fmt(resumenActual.montoCotizado)} cotizado`} icon="💰" color={COLORS.success}/>
              <StatCard label="Días cierre" value={`${resumenActual.diasPromedioCierre}`} sub="Promedio cotización → NV" icon="⏱️" color={COLORS.warning}/>
              <StatCard label="Cotizaciones" value={resumenActual.cotizacionesPeriodo.length} sub={fmt(resumenActual.montoCotizado)} icon="📄" color="#5a8abe"/>
              <StatCard label="Ventas registradas" value={resumenActual.ventasPeriodo.length} sub={`${fmt(resumenActual.montoVendido)} (NV + barranes)`} icon="🧾" color={COLORS.success}/>
              <StatCard label="Ticket cotizado" value={fmt(resumenActual.ticketCotizado)} sub="Promedio por cotización" icon="🧮" color={COLORS.warning}/>
              <StatCard label="Ticket vendido" value={fmt(resumenActual.ticketVendido)} sub="Promedio por venta registrada" icon="💵" color={COLORS.success}/>
              <StatCard label="Clientes" value={resumenActual.clientesVendidos} sub={`${resumenActual.clientesCotizados} cotizados`} icon="👥" color={COLORS.accent}/>
            </div>

            {resumenAnterior && (
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))", gap:12, marginBottom:20 }}>
                <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>Comparación conversión</div>
                  <div style={{ marginTop:8, fontSize:14, color:COLORS.text }}>{rangoResumen.label}: <b style={{ color:COLORS.accent }}>{resumenActual.conversion}%</b></div>
                  <div style={{ fontSize:13, color:COLORS.muted }}>Periodo anterior: {resumenAnterior.conversion}%</div>
                  <div style={{ marginTop:6, fontSize:18, fontWeight:800, color:diferenciaConversionResumen >= 0 ? COLORS.success : COLORS.danger }}>
                    {diferenciaConversionResumen >= 0 ? "+" : ""}{diferenciaConversionResumen} pts
                  </div>
                </div>
                <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>Comparación ventas</div>
                  <div style={{ marginTop:8, fontSize:14, color:COLORS.text }}>{rangoResumen.label}: <b style={{ color:COLORS.success }}>{fmt(resumenActual.montoVendido)}</b></div>
                  <div style={{ fontSize:13, color:COLORS.muted }}>Periodo anterior: {fmt(resumenAnterior.montoVendido)}</div>
                  <div style={{ marginTop:6, fontSize:18, fontWeight:800, color:variacionVentasResumen >= 0 ? COLORS.success : COLORS.danger }}>
                    {variacionVentasResumen >= 0 ? "+" : ""}{variacionVentasResumen}%
                  </div>
                </div>
              </div>
            )}

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))", gap:16, marginBottom:20 }}>
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18 }}>
                <div style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1, marginBottom:12 }}>Conversión por tramo de valor</div>
                <div style={{ display:"grid", gap:10 }}>
                  {resumenActual.conversionPorTramo.map(t => {
                    const ancho = Math.max(3, Math.min(100, t.conversion));
                    return (
                      <div key={t.key}>
                        <div style={{ display:"flex", justifyContent:"space-between", gap:10, fontSize:12, marginBottom:4 }}>
                          <b style={{ color:COLORS.text }}>{t.label}</b>
                          <span style={{ color:t.conversion >= 35 ? COLORS.success : t.conversion >= 20 ? COLORS.warning : COLORS.danger, fontWeight:800 }}>{t.conversion}%</span>
                        </div>
                        <div style={{ height:7, background:COLORS.subtle, borderRadius:6, overflow:"hidden" }}>
                          <div style={{ width:`${ancho}%`, height:"100%", background:t.conversion >= 35 ? COLORS.success : t.conversion >= 20 ? COLORS.warning : COLORS.danger }}/>
                        </div>
                        <div style={{ display:"flex", justifyContent:"space-between", fontSize:10.5, color:COLORS.muted, marginTop:3 }}>
                          <span>{t.vendidas} de {t.cantidad} vendidas</span>
                          <span>{fmt(t.montoVendido)} / {fmt(t.monto)}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18 }}>
                <div style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1, marginBottom:12 }}>Embudo comercial por dinero</div>
                <div style={{ display:"grid", gap:10 }}>
                  {resumenActual.embudoDinero.map(e => {
                    const pct = resumenActual.montoCotizado > 0 ? Math.round((e.monto / resumenActual.montoCotizado) * 100) : 0;
                    const ancho = Math.max(3, Math.min(100, pct));
                    return (
                      <div key={e.key}>
                        <div style={{ display:"flex", justifyContent:"space-between", gap:10, fontSize:12, marginBottom:4 }}>
                          <b style={{ color:e.color }}>{e.label}</b>
                          <span style={{ color:COLORS.text, fontWeight:800 }}>{fmt(e.monto)}</span>
                        </div>
                        <div style={{ height:8, background:COLORS.subtle, borderRadius:6, overflow:"hidden" }}>
                          <div style={{ width:`${ancho}%`, height:"100%", background:e.color }}/>
                        </div>
                        <div style={{ display:"flex", justifyContent:"space-between", fontSize:10.5, color:COLORS.muted, marginTop:3 }}>
                          <span>{e.cantidad} cotizaciones</span>
                          <span>{pct}% del monto cotizado</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {resumenActual.urgentesPeriodo.length>0 && (
              <div style={{ background:"#1e1500", border:`1px solid ${COLORS.warning}`, borderRadius:10, padding:"12px 16px", marginBottom:16 }}>
                <div style={{ fontSize:12, fontWeight:700, color:COLORS.warning, marginBottom:8 }}>⚡ {resumenActual.urgentesPeriodo.length} cotizaciones del período requieren seguimiento</div>
                {resumenActual.urgentesPeriodo.map(q=>(
                  <div key={q.id} style={{ fontSize:12, color:COLORS.text, marginBottom:4, display:"flex", justifyContent:"space-between", alignItems:"center", gap:10 }}>
                    <span><span style={{ color:COLORS.accent }}>#{q.numero}</span> {q.cliente} — {fmt(q.total)}</span>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <span style={{ color:COLORS.muted, fontSize:11 }}>{businessDaysSince(q.fecha)} días háb.</span>
                      <button onClick={()=>setModalCot(q)} style={{ background:COLORS.warning, border:"none", borderRadius:6, padding:"3px 10px", color:"#0f0e0c", fontWeight:700, cursor:"pointer", fontSize:11 }}>+ Nota</button>
                      <button onClick={() => setModalEditarCotizacion(q)} style={{ background:COLORS.accent, border:"none", borderRadius:6, padding:"4px 10px", color:"#0f0e0c", cursor:"pointer", fontSize:12, fontWeight:700 }}>Editar</button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            <div style={{ display:"grid", gridTemplateColumns:"minmax(200px,260px) 1fr", gap:16 }}>
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18, display:"flex", flexDirection:"column", alignItems:"center" }}>
                <span style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1, marginBottom:4 }}>Conversión del período</span>
                <svg width="160" height="92" viewBox="0 0 160 92">
                  <path d={arcPath(startA,startA+sweepA)} stroke={COLORS.subtle} strokeWidth="11" fill="none" strokeLinecap="round"/>
                  <path d={arcPath(startA,startA+sweepA*(resumenActual.conversion/100))} stroke={Number(resumenActual.conversion)>=40?COLORS.success:COLORS.warning} strokeWidth="11" fill="none" strokeLinecap="round"/>
                </svg>
                <div style={{ fontSize:38, fontWeight:800, color:COLORS.accent, fontFamily:"Georgia,serif", lineHeight:1, marginTop:-10 }}>{resumenActual.conversion}%</div>
                <div style={{ fontSize:11, color:COLORS.muted, marginTop:4, textAlign:"center" }}>{resumenActual.vendidasPeriodo.length} vendidas · {resumenActual.vencidasPeriodo.length} vencidas</div>
                <div style={{ marginTop:10, width:"100%", display:"flex", flexDirection:"column", gap:5 }}>
                  {[["✓ Vendidas",resumenActual.vendidasPeriodo.length,COLORS.success],["● Activas",resumenActual.activasPeriodo.length,"#5a8abe"],["⚡ Seguimiento",resumenActual.urgentesPeriodo.length,COLORS.warning],["✕ Vencidas",resumenActual.vencidasPeriodo.length,"#9a7aaa"]].map(([l,c,col])=>(
                    <div key={l} style={{ display:"flex", justifyContent:"space-between" }}>
                      <span style={{ fontSize:11, color:col }}>{l}</span><span style={{ fontSize:11, color:COLORS.muted }}>{c}</span>
                    </div>
                  ))}
                </div>
              </div>

              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18 }}>
                <span style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>Evolución mensual</span>
                <div style={{ marginTop:14, display:"grid", gap:10 }}>
                  {mesesResumenGrafico.map(m => {
                    const anchoCotizado = Math.max(3, Math.round((m.montoCotizado / maxGraficoResumen) * 100));
                    const anchoVendido = Math.max(3, Math.round((m.montoVendido / maxGraficoResumen) * 100));
                    return (
                      <div key={m.key}>
                        <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:4 }}>
                          <b style={{ color:COLORS.text }}>{m.label}</b>
                          <span style={{ color:COLORS.muted }}>Conv. {m.conversion}%</span>
                        </div>
                        <div style={{ display:"grid", gap:3 }}>
                          <div style={{ height:7, background:COLORS.subtle, borderRadius:6, overflow:"hidden" }}><div style={{ width:`${anchoCotizado}%`, height:"100%", background:"#5a8abe" }}/></div>
                          <div style={{ height:7, background:COLORS.subtle, borderRadius:6, overflow:"hidden" }}><div style={{ width:`${anchoVendido}%`, height:"100%", background:COLORS.success }}/></div>
                        </div>
                        <div style={{ display:"flex", justifyContent:"space-between", fontSize:10.5, color:COLORS.muted, marginTop:3 }}>
                          <span>Cotizado {fmt(m.montoCotizado)}</span>
                          <span>Vendido {fmt(m.montoVendido)}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginTop:16 }}>
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18, maxHeight:420, overflowY:"auto" }}>
                <span style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>Estado por Cotización del período</span>
                <div style={{ marginTop:14 }}>
                  {resumenActual.cotizacionesConEstado.slice().sort((a,b)=>{ const o={urgente:0,activa:1,vendida:2,vencida:3}; return o[a.status]-o[b.status]||Number(b.total||0)-Number(a.total||0); }).map(q=>{
                    const nvs=notasVenta.filter(s=>String(s.cotizacion)===String(q.numero));
                    const pct=resumenActual.montoCotizado>0?Number(q.total || 0)/resumenActual.montoCotizado*100:0;
                    const nCount=(seguimiento[q.numero]||[]).length;
                    return (
                      <div key={q.id} style={{ marginBottom:11, opacity:q.status==="vencida"?0.55:1 }}>
                        <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3, gap:8 }}>
                          <span style={{ fontSize:12, color:COLORS.text }}><b style={{ color:COLORS.accent }}>#{q.numero}</b> {q.cliente}</span>
                          <span style={{ fontSize:12, fontWeight:600, color:q.status==="vendida"?COLORS.success:COLORS.muted }}>{fmt(q.total)}</span>
                        </div>
                        <div style={{ height:5, background:COLORS.subtle, borderRadius:3, overflow:"hidden" }}>
                          <div style={{ height:"100%", width:`${pct}%`, background:LEFT_COLOR[q.status], borderRadius:3 }}/>
                        </div>
                        <div style={{ display:"flex", justifyContent:"space-between", marginTop:2, alignItems:"center" }}>
                          <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap" }}>
                            <StatusBadge status={q.status}/>
                            {nvs.length>0 && <span style={{ fontSize:10, color:COLORS.muted }}>→ {nvs.map(n=>`NV#${n.numero}`).join(", ")}</span>}
                            {nCount>0 && <span style={{ fontSize:10, color:COLORS.accent }}>📝{nCount}</span>}
                          </div>
                          <button onClick={()=>setModalCot(q)} style={{ background:"none", border:`1px solid ${COLORS.border}`, borderRadius:6, padding:"2px 8px", color:COLORS.muted, cursor:"pointer", fontSize:11 }}>📝</button>
                        </div>
                      </div>
                    );
                  })}
                  {resumenActual.cotizacionesConEstado.length === 0 && <div style={{ color:COLORS.muted, fontSize:13 }}>Sin cotizaciones en este período.</div>}
                </div>
              </div>

              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18, maxHeight:420, overflowY:"auto" }}>
                <span style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>Top 10 clientes por ventas</span>
                <div style={{ marginTop:14, display:"grid", gap:10 }}>
                  {resumenActual.topClientes.map((c, idx) => (
                    <div key={c.cliente} style={{ display:"grid", gridTemplateColumns:"28px 1fr auto", gap:8, alignItems:"center", paddingBottom:8, borderBottom:`1px solid ${COLORS.border}` }}>
                      <div style={{ color:COLORS.accent, fontWeight:800 }}>#{idx+1}</div>
                      <div>
                        <div style={{ color:COLORS.text, fontSize:13, fontWeight:700 }}>{c.cliente}</div>
                        <div style={{ color:COLORS.muted, fontSize:11 }}>{c.cantidad} NV</div>
                      </div>
                      <div style={{ color:COLORS.success, fontSize:13, fontWeight:800 }}>{fmt(c.total)}</div>
                    </div>
                  ))}
                  {resumenActual.topClientes.length === 0 && <div style={{ color:COLORS.muted, fontSize:13 }}>Sin ventas en este período.</div>}
                </div>
              </div>
            </div>

          {/* SECCIÓN DE ANÁLISIS DE CLIENTE */}
          <div style={{ marginTop:30, borderTop:`2px solid ${COLORS.border}`, paddingTop:20 }}>
            <h2 style={{ color:COLORS.accent, marginTop:0 }}>👥 Análisis de Cliente</h2>
            
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:20 }}>
              <div>
                <label style={{ display:"block", marginBottom:6, fontSize:12, color:COLORS.muted }}>Buscar cliente</label>
                <div style={{ position:"relative" }}>
                  <input
                    type="text"
                    placeholder="Escribe nombre del cliente..."
                    value={busquedaClienteAnalisis}
                    onChange={(e) => setBusquedaClienteAnalisis(e.target.value)}
                    style={{
                      width:"100%",
                      padding:"10px 12px",
                      borderRadius:8,
                      border:`1px solid ${COLORS.border}`,
                      background:COLORS.surface,
                      color:COLORS.text,
                      boxSizing:"border-box",
                      fontSize:13,
                    }}
                  />
                  {busquedaClienteAnalisis && clientesFiltrados.length > 0 && (
                    <div style={{
                      position:"absolute",
                      top:"100%",
                      left:0,
                      right:0,
                      background:COLORS.surface,
                      border:`1px solid ${COLORS.border}`,
                      borderTop:"none",
                      borderRadius:"0 0 8px 8px",
                      maxHeight:200,
                      overflowY:"auto",
                      zIndex:100,
                    }}>
                      {clientesFiltrados.map((c) => (
                        <div
                          key={c}
                          onClick={() => { setClienteSeleccionado(c); setBusquedaClienteAnalisis(""); }}
                          style={{
                            padding:"10px 12px",
                            cursor:"pointer",
                            borderBottom:`1px solid ${COLORS.border}`,
                            fontSize:13,
                          }}
                          onMouseEnter={(e) => (e.target.style.background = COLORS.card)}
                          onMouseLeave={(e) => (e.target.style.background = "transparent")}
                        >
                          {c}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              <div>
                <label style={{ display:"block", marginBottom:6, fontSize:12, color:COLORS.muted }}>Período</label>
                <select
                  value={filtroFechasAnalisis}
                  onChange={(e) => setFiltroFechasAnalisis(e.target.value)}
                  style={{
                    width:"100%",
                    padding:"10px 12px",
                    borderRadius:8,
                    border:`1px solid ${COLORS.border}`,
                    background:COLORS.surface,
                    color:COLORS.text,
                    boxSizing:"border-box",
                    fontSize:13,
                  }}
                >
                  <option value="mes_actual">Mes actual</option>
                  <option value="ultimos_3_meses">Últimos 3 meses</option>
                  <option value="ultimo_año">Último año</option>
                  <option value="año_actual">Año actual</option>
                </select>
              </div>
            </div>

            {clienteSeleccionado && analisisClienteSeleccionado && (
              <div>
                <div style={{ marginBottom:20, padding:16, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                    <h3 style={{ margin:0, color:COLORS.accent, fontSize:16 }}>{clienteSeleccionado}</h3>
                    <button
                      onClick={() => setClienteSeleccionado(null)}
                      style={{ background:"none", border:"none", color:COLORS.muted, cursor:"pointer", fontSize:18 }}
                    >✕</button>
                  </div>

                  <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))", gap:10, marginBottom:16 }}>
                    <div style={{ background:COLORS.surface, padding:12, borderRadius:8, border:`1px solid ${COLORS.border}` }}>
                      <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>Total Cotizado</div>
                      <div style={{ fontSize:16, fontWeight:700, color:COLORS.accent }}>{fmt(analisisClienteSeleccionado.totalCotizado)}</div>
                    </div>
                    <div style={{ background:COLORS.surface, padding:12, borderRadius:8, border:`1px solid ${COLORS.border}` }}>
                      <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>Cotizaciones</div>
                      <div style={{ fontSize:16, fontWeight:700, color:"#5a8abe" }}>{analisisClienteSeleccionado.nroCotizaciones}</div>
                    </div>
                    <div style={{ background:COLORS.surface, padding:12, borderRadius:8, border:`1px solid ${COLORS.border}` }}>
                      <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>Total Vendido</div>
                      <div style={{ fontSize:16, fontWeight:700, color:COLORS.success }}>{fmt(analisisClienteSeleccionado.totalVendido)}</div>
                    </div>
                    <div style={{ background:COLORS.surface, padding:12, borderRadius:8, border:`1px solid ${COLORS.border}` }}>
                      <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>Notas de Venta</div>
                      <div style={{ fontSize:16, fontWeight:700, color:COLORS.warning }}>{analisisClienteSeleccionado.nroNotas}</div>
                    </div>
                    <div style={{ background:COLORS.surface, padding:12, borderRadius:8, border:`1px solid ${COLORS.border}` }}>
                      <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>Conversión</div>
                      <div style={{ fontSize:16, fontWeight:700, color:COLORS.accent }}>{analisisClienteSeleccionado.tasaConversion}%</div>
                    </div>
                    <div style={{ background:COLORS.surface, padding:12, borderRadius:8, border:`1px solid ${COLORS.border}` }}>
                      <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>Frecuencia</div>
                      <div style={{ fontSize:13, fontWeight:700, color:"#8a8270" }}>{analisisClienteSeleccionado.frecuenciaPromedio}</div>
                    </div>
                  </div>

                  {analisisClienteSeleccionado.ultimaCompra && (
                    <div style={{ fontSize:12, color:COLORS.muted, padding:10, background:COLORS.bg, borderRadius:6 }}>
                      📅 Última compra: <b>{new Date(analisisClienteSeleccionado.ultimaCompra.fecha).toLocaleDateString("es-CL")}</b> por {fmt(analisisClienteSeleccionado.ultimaCompra.monto)}
                    </div>
                  )}
                </div>

                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))", gap:14 }}>
                  <div style={{ padding:14, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10 }}>
                    <h4 style={{ margin:"0 0 12px", color:COLORS.accent, fontSize:13 }}>🛋️ Tipos de Cubierta</h4>
                    {analisisClienteSeleccionado.topCubiertas.length === 0 ? <div style={{ fontSize:12, color:COLORS.muted }}>Sin datos</div> : analisisClienteSeleccionado.topCubiertas.map((item, i) => <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:6 }}><span style={{ color:COLORS.text }}>{item.tipo}</span><span style={{ fontWeight:700, color:COLORS.accent }}>{item.cantidad}x</span></div>)}
                  </div>
                  <div style={{ padding:14, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10 }}>
                    <h4 style={{ margin:"0 0 12px", color:COLORS.accent, fontSize:13 }}>📏 Medidas</h4>
                    {analisisClienteSeleccionado.topMedidas.length === 0 ? <div style={{ fontSize:12, color:COLORS.muted }}>Sin datos</div> : analisisClienteSeleccionado.topMedidas.map((item, i) => <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:6 }}><span style={{ color:COLORS.text }}>{item.medida}</span><span style={{ fontWeight:700, color:COLORS.warning }}>{item.cantidad}x</span></div>)}
                  </div>
                  <div style={{ padding:14, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10 }}>
                    <h4 style={{ margin:"0 0 12px", color:COLORS.accent, fontSize:13 }}>🎨 Colores</h4>
                    {analisisClienteSeleccionado.topColores.length === 0 ? <div style={{ fontSize:12, color:COLORS.muted }}>Sin datos</div> : analisisClienteSeleccionado.topColores.map((item, i) => <div key={i} style={{ display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:6 }}><span style={{ color:COLORS.text }}>{item.color}</span><span style={{ fontWeight:700, color:"#8a7c5a" }}>{item.cantidad}x</span></div>)}
                  </div>
                </div>
              </div>
            )}

            {!clienteSeleccionado && (
              <div style={{ padding:20, background:COLORS.card, border:`1px dashed ${COLORS.border}`, borderRadius:10, textAlign:"center", color:COLORS.muted, fontSize:14 }}>
                Busca un cliente para ver su análisis
              </div>
            )}
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
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setModalEditarCotizacion(q);
                          }}
                          style={{
                            background: COLORS.accent,
                            border: "none",
                            borderRadius: 6,
                            padding: "4px 10px",
                            color: "#0f0e0c",
                            cursor: "pointer",
                            fontSize: 12,
                            fontWeight: 700
                          }}
                        >
                          Editar
                        </button>
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
    {s.cotizacion ? (
  <span
    title={
      conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
        ? `Esta cotización está asociada a ${conteoCotizacionesEnNV[String(s.cotizacion || "").trim()]} notas/barranes`
        : ""
    }
    style={{
      marginLeft: 8,
      fontSize: 11,
      color:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? "#fff"
          : COLORS.muted,
      background:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? COLORS.danger
          : COLORS.subtle,
      borderRadius: 4,
      padding: "2px 7px",
      border:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? `1px solid ${COLORS.danger}`
          : "none",
      fontWeight:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? 800
          : 400
    }}
  >
    {conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
      ? `⚠ COT#${s.cotizacion} duplicada`
      : `← COT#${s.cotizacion}`}
  </span>
) : (
  <span style={{ marginLeft:8, fontSize:11, color:COLORS.warning, background:"#2a1f0a", borderRadius:4, padding:"2px 7px", border:`1px solid ${COLORS.warning}` }}>
    sin cotización
  </span>
)}
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
      background:estadoProduccionDesdeNota(s).color + "22",
      color:estadoProduccionDesdeNota(s).color,
      border:`1px solid ${estadoProduccionDesdeNota(s).color}`
    }}>
      {estadoProduccionDesdeNota(s).texto}
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
      onClick={(e) => {
        e.stopPropagation();
        setModalEditarNV(s);
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
  onClick={async (e) => {
    e.stopPropagation();
    const confirmar = await confirmDialog(`¿Eliminar NV ${s.numero}?`);

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
{tab==="barranes" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:14 }}>
              <h2 style={{ margin:0, fontFamily:"Georgia,serif", color:COLORS.success, fontSize:17 }}>Barranes {mesSeleccionadoTexto}</h2>
              <input
                type="month"
                value={mesFiltro}
                onChange={(e) => setMesFiltro(e.target.value)}
                style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }}
              />
            </div>
            <div style={{ marginBottom:14, background:COLORS.subtle, borderRadius:10, padding:"10px 16px", display:"flex", gap:24 }}>
              <div><span style={{ fontSize:11, color:COLORS.muted }}>TOTAL BARRANES</span><div style={{ fontSize:18, fontWeight:700, color:COLORS.success }}>{fmt(totalBarranesFiltrado)}</div></div>
              <div><span style={{ fontSize:11, color:COLORS.muted }}>BARRANES</span><div style={{ fontSize:18, fontWeight:700, color:COLORS.success }}>{filteredBarranes.length}</div></div>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
              {filteredBarranes.map(s=>(
                <div
                key={s.id}
                onClick={() => setModalNV(s)}
                style={{background:COLORS.card, border:`1px solid ${s.cotizacion?"#2d5040":COLORS.border}`, borderLeft:`4px solid ${s.cotizacion?COLORS.success:COLORS.warning}`, borderRadius:10, padding:"11px 14px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                  <div>
  <div>
    <span style={{ fontWeight:700, color:COLORS.success, marginRight:8 }}>Barrán #{s.numero}</span>
    <span style={{ color:COLORS.text }}>{s.cliente}</span>
    {s.cotizacion ? (
  <span
    title={
      conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
        ? `Esta cotización está asociada a ${conteoCotizacionesEnNV[String(s.cotizacion || "").trim()]} notas/barranes`
        : ""
    }
    style={{
      marginLeft: 8,
      fontSize: 11,
      color:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? "#fff"
          : COLORS.muted,
      background:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? COLORS.danger
          : COLORS.subtle,
      borderRadius: 4,
      padding: "2px 7px",
      border:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? `1px solid ${COLORS.danger}`
          : "none",
      fontWeight:
        conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
          ? 800
          : 400
    }}
  >
    {conteoCotizacionesEnNV[String(s.cotizacion || "").trim()] > 1
      ? `⚠ COT#${s.cotizacion} duplicada`
      : `← COT#${s.cotizacion}`}
  </span>
) : (
  <span style={{ marginLeft:8, fontSize:11, color:COLORS.warning, background:"#2a1f0a", borderRadius:4, padding:"2px 7px", border:`1px solid ${COLORS.warning}` }}>
    barrán interno
  </span>
)}
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
      background:estadoProduccionDesdeNota(s).color + "22",
      color:estadoProduccionDesdeNota(s).color,
      border:`1px solid ${estadoProduccionDesdeNota(s).color}`
    }}>
      {estadoProduccionDesdeNota(s).texto}
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
      onClick={(e) => {
        e.stopPropagation();
        setModalEditarNV(s);
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
      Editar
    </button>
  )}
  <button
  onClick={async (e) => {
    e.stopPropagation();
    const confirmar = await confirmDialog(`¿Eliminar Barrán ${s.numero}?`);

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

    <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:14 }}>
      <span style={{ fontSize:12, color:COLORS.muted, fontWeight:700 }}>Ordenar por:</span>
      <button
        onClick={() => setOrdenProduccion("correlativo")}
        style={{
          border:`1px solid ${ordenProduccion === "correlativo" ? COLORS.success : COLORS.border}`,
          background:ordenProduccion === "correlativo" ? COLORS.success : COLORS.card,
          color:ordenProduccion === "correlativo" ? "#fff" : COLORS.text,
          borderRadius:8,
          padding:"7px 10px",
          cursor:"pointer",
          fontSize:12,
          fontWeight:700
        }}
      >
        Nº mayor a menor
      </button>
      <button
        onClick={() => setOrdenProduccion("fecha_entrega")}
        style={{
          border:`1px solid ${ordenProduccion === "fecha_entrega" ? COLORS.success : COLORS.border}`,
          background:ordenProduccion === "fecha_entrega" ? COLORS.success : COLORS.card,
          color:ordenProduccion === "fecha_entrega" ? "#fff" : COLORS.text,
          borderRadius:8,
          padding:"7px 10px",
          cursor:"pointer",
          fontSize:12,
          fontWeight:700
        }}
      >
        Fecha de entrega
      </button>
    </div>

    <div style={{ marginBottom:18 }}>
      <h3 style={{ margin:"0 0 10px", color:COLORS.success, fontSize:15 }}>
        Notas de Venta ({trabajosProduccionNV.length})
      </h3>

      {trabajosProduccionNV.length === 0 ? (
        <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, color:COLORS.muted }}>
          No hay notas de venta para mostrar con este filtro.
        </div>
      ) : (
        <div style={{ display:"grid", gap:12 }}>
          {trabajosProduccionNV.map(renderTarjetaProduccion)}
        </div>
      )}
    </div>

    <div style={{ marginTop:20 }}>
      <h3 style={{ margin:"0 0 10px", color:"#8b5cf6", fontSize:15 }}>
        Barranes ({trabajosProduccionBarranes.length})
      </h3>

      {trabajosProduccionBarranes.length === 0 ? (
        <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, color:COLORS.muted }}>
          No hay barranes para mostrar con este filtro.
        </div>
      ) : (
        <div style={{ display:"grid", gap:12 }}>
          {trabajosProduccionBarranes.map(renderTarjetaProduccion)}
        </div>
      )}
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
    multiple
    onChange={importarOrdenCompraExcel}
    style={{ display:"none" }}
  />
</label>
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
  📄 Importar Factura XML
  <input
    type="file"
    accept=".xml,text/xml,application/xml"
    onChange={importarFacturaXml}
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
      display:"grid",
      gridTemplateColumns:window.innerWidth <= 900 ? "1fr" : "minmax(0, 1fr) minmax(260px, 320px)",
      gap:16,
      alignItems:"start"
    }}>
      <div>
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

      <aside style={{
        background:COLORS.card,
        border:`1px solid ${COLORS.border}`,
        borderRadius:12,
        padding:14,
        position:window.innerWidth <= 900 ? "static" : "sticky",
        top:12,
        maxHeight:"calc(100vh - 120px)",
        overflowY:"auto"
      }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:10, marginBottom:10 }}>
          <div>
            <p style={{ margin:"0 0 4px", color:COLORS.text, fontWeight:700 }}>
              OC importadas
            </p>
            <p style={{ margin:0, color:COLORS.muted, fontSize:12 }}>
              Ordenadas de mayor a menor
            </p>
          </div>
          <span style={{ color:COLORS.accent, fontWeight:800 }}>
            {ordenesCompraImportadas.length}
          </span>
        </div>

        {ordenesCompraImportadas.length === 0 ? (
          <div style={{ color:COLORS.muted, fontSize:13 }}>
            Todavía no hay OC registradas.
          </div>
        ) : (
          <div style={{ display:"grid", gap:8 }}>
            {ordenesCompraImportadas.map((doc, idx) => (
              <div
                key={doc.id || doc.clave || idx}
                style={{
                  display:"grid",
                  gridTemplateColumns:"70px 1fr",
                  gap:8,
                  alignItems:"center",
                  padding:"8px 0",
                  borderBottom:idx === ordenesCompraImportadas.length - 1 ? "none" : `1px solid ${COLORS.border}`
                }}
              >
                <b style={{ color:COLORS.accent, fontSize:13 }}>
                  {doc.documento || doc.folio || "S/N"}
                </b>
                <span style={{ color:COLORS.muted, fontSize:11, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>
                  {fmtDate(String(doc.fecha || "").split("T")[0])}
                </span>
              </div>
            ))}
          </div>
        )}
      </aside>
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



        {tab==="calendario" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, flexWrap:"wrap", marginBottom:16 }}>
              <div>
                <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>📅 Calendario</h2>
                <p style={{ margin:0, fontSize:12, color:COLORS.muted }}>Entregas de NV/Barranes y pagos comprometidos del mes.</p>
              </div>
              <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                <button onClick={() => cambiarMesCalendario(-1)} style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", cursor:"pointer" }}>← Mes anterior</button>
                <input type="month" value={mesCalendario} onChange={e=>setMesCalendario(e.target.value)} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 12px" }}/>
                <button onClick={() => cambiarMesCalendario(1)} style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", cursor:"pointer" }}>Mes siguiente →</button>
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12, marginBottom:14 }}>
              <StatCard label="Entregas del mes" value={entregasPendientesMes.length} icon="🚚" color="#60a5fa" />
              <StatCard label="Pagos pendientes" value={pagosPendientesMes.length} icon="💸" color={COLORS.warning} />
              <StatCard label="Comprometido por pagar" value={fmt(totalPagosPendientesMes)} icon="📌" color={COLORS.danger} />
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(7, minmax(130px, 1fr))", gap:6, overflowX:"auto" }}>
              {["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"].map(d => (
                <div key={d} style={{ minWidth:130, color:COLORS.muted, fontSize:12, fontWeight:700, padding:"6px 8px", textAlign:"center" }}>{d}</div>
              ))}
              {Array.from({ length:celdasVaciasInicio }).map((_, i) => <div key={`empty-${i}`} style={{ minWidth:130, minHeight:95 }} />)}
              {Array.from({ length:diasMesCalendario }).map((_, idx) => {
                const dia = idx + 1;
                const fechaDia = `${mesCalendario}-${String(dia).padStart(2, "0")}`;
                const eventosDia = eventosDelMes.filter(e => e.fecha === fechaDia);
                return (
                  <div key={fechaDia} style={{ minWidth:130, minHeight:115, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:8 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                      <b style={{ color:COLORS.accent }}>{dia}</b>
                      {eventosDia.length > 0 && <span style={{ fontSize:11, color:COLORS.muted }}>{eventosDia.length}</span>}
                    </div>
                    <div style={{ display:"grid", gap:5 }}>
                      {eventosDia.map((ev, i) => (
                        <button key={`${fechaDia}-${i}`} onClick={ev.onClick} style={{ textAlign:"left", background:ev.color + "22", border:`1px solid ${ev.color}`, color:COLORS.text, borderRadius:7, padding:"5px 6px", cursor:"pointer", fontSize:11 }}>
                          <b style={{ color:ev.color }}>{ev.titulo}</b><br />
                          <span style={{ color:COLORS.muted }}>{ev.subtitulo}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab==="cuentas_pagar" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-end", flexWrap:"wrap", marginBottom:16 }}>
              <div>
                <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>💸 Cuentas por pagar</h2>
                <p style={{ margin:0, fontSize:12, color:COLORS.muted }}>Deudas, cuotas, proveedores, imposiciones y compromisos visibles en el Calendario.</p>
              </div>
              <button onClick={() => setModalCuentaPagar(true)} style={{ background:COLORS.accent, color:"#111", border:"none", borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Nueva cuenta</button>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12, marginBottom:14 }}>
              <StatCard label="Pendiente total" value={fmt(cuentasPagar.reduce((sum, c) => sum + saldoCuentaPagar(c), 0))} icon="📌" color={COLORS.danger} />
              <StatCard label="Por pagar este mes" value={fmt(totalPagosPendientesMes)} icon="📅" color={COLORS.warning} />
              <StatCard label="Cuentas activas" value={cuentasPagar.filter(c => saldoCuentaPagar(c) > 0).length} icon="💸" color={COLORS.accent} />
            </div>

            <div style={{ display:"grid", gap:10 }}>
              {cuentasPagar.length === 0 ? (
                <div style={{ color:COLORS.muted, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:14 }}>No hay cuentas por pagar registradas.</div>
              ) : cuentasPagar.map(c => {
                const abonos = obtenerAbonosCuenta(c.id);
                const abonado = abonos.reduce((sum, a) => sum + Number(a.monto || 0), 0);
                const saldo = saldoCuentaPagar(c);
                return (
                  <div key={c.id} style={{ background:COLORS.card, border:`1px solid ${saldo === 0 ? COLORS.success : COLORS.border}`, borderRadius:12, padding:14 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", gap:10, flexWrap:"wrap" }}>
                      <div>
                        <b style={{ color:COLORS.accent }}>{c.nombre}</b>
                        <div style={{ fontSize:12, color:COLORS.muted }}>{c.tipo} · Vence {fmtDate(c.fecha_vencimiento)} · Cuota {c.cuota_actual || 1}/{c.cuotas_totales || 1}</div>
                        {c.detalle && <div style={{ fontSize:12, marginTop:5 }}>{c.detalle}</div>}
                      </div>
                      <div style={{ textAlign:"right" }}>
                        <div><b>Total:</b> {fmt(c.monto_total)}</div>
                        <div><b>Abonado:</b> {fmt(abonado)}</div>
                        <div style={{ color:saldo === 0 ? COLORS.success : COLORS.warning, fontWeight:800 }}><b>Saldo:</b> {fmt(saldo)}</div>
                      </div>
                    </div>
                    <div style={{ marginTop:10, display:"flex", gap:8, flexWrap:"wrap" }}>
                      <button onClick={() => setModalAbonoCuentaPagar(c)} disabled={saldo === 0} style={{ background:saldo === 0 ? COLORS.subtle : COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"8px 10px", fontWeight:700, cursor:saldo === 0 ? "not-allowed" : "pointer" }}>Registrar abono</button>
                    </div>
                    {abonos.length > 0 && (
                      <div style={{ marginTop:10, fontSize:12, color:COLORS.muted }}>
                        Historial: {abonos.map(a => `${fmtDate(a.fecha)} ${fmt(a.monto)} ${a.medio_pago || ""}`).join(" · ")}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab==="rrhh" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-end", flexWrap:"wrap", marginBottom:16 }}>
              <div>
                <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>👷 Recursos Humanos</h2>
                <p style={{ margin:0, fontSize:12, color:COLORS.muted }}>Fichas básicas de trabajadores. Los documentos se guardan fuera de la app para no gastar almacenamiento de Supabase.</p>
              </div>
              <button onClick={() => setModalTrabajador(true)} style={{ background:COLORS.accent, color:"#111", border:"none", borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Trabajador</button>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))", gap:12 }}>
              {trabajadores.length === 0 ? (
                <div style={{ color:COLORS.muted, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:14 }}>No hay trabajadores registrados.</div>
              ) : trabajadores.map(t => {
                const docs = documentosTrabajadores.filter(d => String(d.trabajador_id) === String(t.id));
                return (
                  <div key={t.id} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", gap:10 }}>
                      <div>
                        <b style={{ color:COLORS.accent }}>{t.nombre}</b>
                        <div style={{ fontSize:12, color:COLORS.muted }}>{t.rut || "Sin RUT"} · {t.cargo || "Sin cargo"}</div>
                        <div style={{ fontSize:12, color:COLORS.muted }}>Ingreso: {t.fecha_ingreso ? fmtDate(t.fecha_ingreso) : "-"} · Estado: {t.estado || "activo"}</div>
                      </div>
                      <div style={{ fontWeight:800, color:t.estado === "finiquitado" ? COLORS.danger : COLORS.success }}>{t.estado || "activo"}</div>
                    </div>

                    <div style={{ marginTop:10, fontSize:12, color:COLORS.muted }}>
                      Documentos: guardar en carpeta externa / Drive / correo. La app no sube archivos para ahorrar Storage.
                    </div>

                    <div style={{ marginTop:12, display:"grid", gap:6 }}>
                      {docs.length === 0 ? <div style={{ fontSize:12, color:COLORS.muted }}>Sin documentos.</div> : docs.map(d => (
                        <a key={d.id} href={d.url_publica} target="_blank" rel="noreferrer" style={{ color:COLORS.text, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"7px 8px", fontSize:12, textDecoration:"none" }}>
                          📄 {d.tipo} · {d.nombre_archivo}
                        </a>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab==="cartola" && (
          <section style={{ maxWidth:1200, margin:"0 auto" }}>
            <h2 style={{ color:COLORS.accent, fontFamily:"Georgia,serif", marginBottom:6 }}>🏦 Cartola bancaria</h2>
            <p style={{ color:COLORS.muted, marginTop:0, fontSize:13 }}>Importa Excel/CSV del banco, revisa cada movimiento y conviértelo en abono, pago o asiento contable.</p>
            <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, marginBottom:14 }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-end", gap:10, flexWrap:"wrap", marginBottom:12 }}>
                <div>
                  <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent }}>🏦 Cartola bancaria</div>
                  <div style={{ color:COLORS.muted, fontSize:12 }}>Importa Excel/CSV del banco, revisa movimientos y conviértelos en abonos, pagos o asientos.</div>
                </div>
                <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap" }}>
  <button onClick={() => setMesCartola(cambiarMesSimple(mesCartola, -1))} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px", cursor:"pointer" }}>
    ←
  </button>

  <input type="month" value={mesCartola} onChange={e=>setMesCartola(e.target.value)} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }} />

  <button onClick={() => setMesCartola(cambiarMesSimple(mesCartola, 1))} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px", cursor:"pointer" }}>
    →
  </button>

  <button onClick={() => setMesCartola(mesActual)} style={{ background:COLORS.accent, border:"none", color:"#0f0e0c", borderRadius:8, padding:"8px 10px", fontWeight:700, cursor:"pointer" }}>
    Mes actual
  </button>
</div>
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, marginBottom:12 }}>
                <StatCard label="Ingresos cartola" value={fmt(totalIngresosCartola)} icon="⬆️" color={COLORS.success}/>
                <StatCard label="Egresos cartola" value={fmt(totalEgresosCartola)} icon="⬇️" color={COLORS.danger}/>
                <StatCard label="Por revisar" value={cartolaPorRevisar.length} icon="🔎" color={cartolaPorRevisar.length ? COLORS.warning : COLORS.success}/>
              </div>

              <div style={{ marginBottom:12, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:12 }}>
                <div style={{ fontSize:12, color:COLORS.accent, fontWeight:800, marginBottom:8 }}>Resumen por categoría del mes</div>
                {resumenCartolaCategorias.length === 0 ? (
                  <div style={{ color:COLORS.muted, fontSize:12 }}>Sin movimientos clasificados.</div>
                ) : (
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:8 }}>
                    {resumenCartolaCategorias.slice(0, 8).map(c => (
                      <div key={c.categoria} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:10 }}>
                        <div style={{ fontSize:11, color:COLORS.muted }}>{c.categoria}</div>
                        <div style={{ fontSize:14, color:c.ingresos > c.egresos ? COLORS.success : COLORS.danger, fontWeight:800 }}>
                          {fmt(c.ingresos > 0 ? c.ingresos : c.egresos)}
                        </div>
                        <div style={{ fontSize:10, color:COLORS.muted }}>{c.movimientos} mov.</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div style={{ display:"flex", justifyContent:"space-between", gap:10, flexWrap:"wrap", alignItems:"center", marginBottom:10 }}>
                <input value={busquedaCartola} onChange={e=>setBusquedaCartola(e.target.value)} placeholder="Buscar en cartola..." style={{ minWidth:220, flex:"0 1 320px", background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }} />
                <span style={{ color:COLORS.muted, fontSize:12 }}>{cartolaFiltrada.length} movimientos</span>
              </div>

              {cartolaFiltrada.length === 0 ? (
                <div style={{ color:COLORS.muted, fontSize:12 }}>Todavía no hay cartola importada para este mes.</div>
              ) : (
                <div style={{ overflowX:"auto", maxHeight:360, overflowY:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, minWidth:900 }}>
                    <thead>
                      <tr style={{ color:COLORS.muted, textAlign:"left" }}>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Fecha</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Descripción</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Cargo</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Abono</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Categoría</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Estado</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Detalle / Acción</th>
                      </tr>
                    </thead>
                    <tbody>
                      {cartolaFiltrada.slice(0, 80).map(m => (
                        <tr key={m.id}>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{fmtDate(m.fecha)}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, color:COLORS.text, fontWeight:700 }}>{m.descripcion}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.danger }}>{Number(m.cargo || 0) ? fmt(m.cargo) : "-"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.success }}>{Number(m.abono || 0) ? fmt(m.abono) : "-"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>
                            <span style={{
                              display:"inline-block",
                              padding:"3px 8px",
                              borderRadius:999,
                              background:(m.categoria && m.categoria !== "Por revisar") ? "rgba(90,158,111,.16)" : "rgba(200,148,58,.14)",
                              color:(m.categoria && m.categoria !== "Por revisar") ? COLORS.success : COLORS.warning,
                              fontWeight:700
                            }}>
                              {m.categoria || "Por revisar"}
                            </span>
                          </td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, color:m.estado === "conciliado" ? COLORS.success : COLORS.warning }}>{m.estado === "conciliado" ? "Clasificado" : "Por revisar"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>
                            <div style={{ display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
                              <button onClick={() => abrirClasificacionCartola(m)} style={{ background:COLORS.accent, color:"#111", border:"none", borderRadius:7, padding:"6px 9px", cursor:"pointer", fontSize:11, fontWeight:800 }}>Clasificar</button>
                              {Number(m.abono || 0) > 0 && <button onClick={() => { setModalClasificarCartola(m); setCategoriaCartolaSeleccionada("Abono CxC"); setBusquedaCategoriaCartola(""); setBusquedaCxcCartola(""); setCxcSeleccionadaCartola(null); setTipoAbonoCartola("abono"); }} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11 }}>Abono CxC</button>}
                              <span style={{ color:COLORS.muted, fontSize:11 }}>{m.nota_venta_id ? `Doc asociado #${m.nota_venta_id}` : (m.observacion || "")}</span>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </section>
        )}

        
        {tab==="cxc" && (
          <section style={{ maxWidth:1200, margin:"0 auto" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-end", gap:12, flexWrap:"wrap", marginBottom:14 }}>
              <div>
                <h2 style={{ color:COLORS.accent, fontFamily:"Georgia,serif", margin:"0 0 6px" }}>💰 Cuentas por cobrar</h2>
                <p style={{ color:COLORS.muted, margin:0, fontSize:13 }}>
                  Saldos conectados con Notas de Venta, Barranes, abonos y Producción.
                </p>
              </div>
              <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                <button onClick={() => setMesCxc(cambiarMesSimple(mesCxc, -1))} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px", cursor:"pointer" }}>
                  ←
                </button>

                <input
                  type="month"
                  value={mesCxc}
                  onChange={e=>setMesCxc(e.target.value)}
                  style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }}
                />

                <button onClick={() => setMesCxc(cambiarMesSimple(mesCxc, 1))} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px", cursor:"pointer" }}>
                  →
                </button>

                <button onClick={() => setMesCxc(mesActual)} style={{ background:COLORS.accent, border:"none", color:"#0f0e0c", borderRadius:8, padding:"8px 10px", fontWeight:700, cursor:"pointer" }}>
                  Mes actual
                </button>

                <input
                  value={busquedaCxcCartola}
                  onChange={e=>setBusquedaCxcCartola(e.target.value)}
                  placeholder="Buscar cliente, NV, barrán, detalle..."
                  style={{ minWidth:260, background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 11px" }}
                />
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12, marginBottom:16 }}>
              <StatCard label="Total por cobrar" value={fmt(totalCxcNotas)} sub={`${detalleCxcNotas.length} documentos`} icon="💰" color={COLORS.accent}/>
              <StatCard label="Saldos pendientes" value={fmt(totalCxcSaldosPendientes)} sub={`${cxcSaldosPendientes.length} documentos`} icon="🟡" color={COLORS.warning}/>
              <StatCard label="Entregadas por confirmar" value={fmt(totalCxcEntregadasPorConfirmar)} sub={`${cxcEntregadasPorConfirmar.length} revisar`} icon="🔴" color={COLORS.danger}/>
              <StatCard label="Crédito autorizado" value={fmt(totalCxcCreditoAutorizado)} sub={`${cxcCreditoAutorizado.length} documentos`} icon="🟣" color="#a78bfa"/>
            </div>

            {cxcCreditoAutorizado.length > 0 && (
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, marginBottom:16 }}>
                <div style={{ fontSize:12, color:COLORS.accent, fontWeight:800, marginBottom:8 }}>🟣 Desglose crédito autorizado</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:8 }}>
                  {Object.values(cxcCreditoAutorizado.reduce((acc, item) => {
                    const cliente = item.cliente || "Sin cliente";
                    if (!acc[cliente]) acc[cliente] = { cliente, saldo:0, docs:0 };
                    acc[cliente].saldo += Number(item.saldo || 0);
                    acc[cliente].docs += 1;
                    return acc;
                  }, {})).map(c => (
                    <div key={c.cliente} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:10 }}>
                      <div style={{ fontWeight:800, color:COLORS.text, fontSize:13 }}>{c.cliente}</div>
                      <div style={{ color:"#a78bfa", fontWeight:800 }}>{fmt(c.saldo)}</div>
                      <div style={{ color:COLORS.muted, fontSize:11 }}>{c.docs} documento(s)</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, overflow:"hidden" }}>
              {detalleCxcFiltrado.length === 0 ? (
                <div style={{ padding:18, color:COLORS.muted }}>No hay CxC pendientes con los filtros actuales.</div>
              ) : (
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, minWidth:980 }}>
                    <thead>
                      <tr style={{ color:COLORS.muted, textAlign:"left", background:COLORS.surface }}>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>Estado</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>Documento</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>Cliente / Detalle</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Total</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Abonado</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Saldo</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>Producción</th>
                        <th style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>Acciones</th>
                      </tr>
                    </thead>
                    <tbody>
                      {detalleCxcFiltrado.map(item => (
                        <tr key={item.id}>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>
                            <span style={{
                              padding:"4px 8px",
                              borderRadius:999,
                              background:`${item.estado.color}22`,
                              color:item.estado.color,
                              fontWeight:800,
                              whiteSpace:"nowrap"
                            }}>
                              {item.estado.icon} {item.estado.label}
                            </span>
                            {item.estado.key === "entregada_por_confirmar" && (
                              <div style={{ marginTop:5, color:COLORS.danger, fontSize:10.5, lineHeight:1.25 }}>
                                Esta venta fue marcada como entregada. Confirmar si el saldo fue pagado.
                              </div>
                            )}
                          </td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.accent, fontWeight:800 }}>{item.tipo} #{item.numero}</td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>
                            <div style={{ color:COLORS.text, fontWeight:800 }}>{item.cliente}</div>
                            <div style={{ color:COLORS.muted, fontSize:11 }}>{item.fecha ? fmtDate(item.fecha) : ""} {item.detalleTexto ? `· ${item.detalleTexto}` : ""}</div>
                          </td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>{fmt(item.total)}</td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.success }}>{fmt(item.abonado)}</td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:item.estado.color, fontWeight:900 }}>{fmt(item.saldo)}</td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>
                            <div style={{ color:COLORS.text }}>{item.proceso || "en espera"}</div>
                            {item.fechaEntrega && <div style={{ color:COLORS.muted, fontSize:11 }}>Entrega: {fmtDate(item.fechaEntrega)}</div>}
                          </td>
                          <td style={{ padding:10, borderBottom:`1px solid ${COLORS.border}` }}>
                            <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                              <button
                                onClick={() => setModalAbonoCxc(item)}
                                style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11, fontWeight:800 }}
                              >
                                Registrar abono
                              </button>
                              <button
                                onClick={() => toggleCreditoAutorizadoCliente(item.cliente)}
                                style={{ background:clienteTieneCreditoAutorizado(item.cliente) ? "#a78bfa" : COLORS.surface, color:clienteTieneCreditoAutorizado(item.cliente) ? "#111" : COLORS.text, border:`1px solid ${clienteTieneCreditoAutorizado(item.cliente) ? "#a78bfa" : COLORS.border}`, borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11, fontWeight:800 }}
                              >
                                {clienteTieneCreditoAutorizado(item.cliente) ? "Quitar crédito" : "Crédito autorizado"}
                              </button>
                              <button
                                onClick={() => setModalNV(item.nota)}
                                style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11 }}
                              >
                                Ver detalle
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </section>
        )}

        {tab==="contabilidad" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-end", flexWrap:"wrap", marginBottom:16 }}>
              <div>
                <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>📚 Contabilidad</h2>
                <p style={{ margin:0, fontSize:12, color:COLORS.muted }}>Libro diario simple con Debe, Haber, IVA CF/DF, Banco, Caja y CxC Clientes.</p>
              </div>
              <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap" }}>
  <button onClick={() => setMesContabilidad(cambiarMesSimple(mesContabilidad, -1))} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px", cursor:"pointer" }}>
    ←
  </button>

  <input type="month" value={mesContabilidad} onChange={e=>setMesContabilidad(e.target.value)} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }} />

  <button onClick={() => setMesContabilidad(cambiarMesSimple(mesContabilidad, 1))} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px", cursor:"pointer" }}>
    →
  </button>

  <button onClick={() => setMesContabilidad(mesActual)} style={{ background:COLORS.accent, border:"none", color:"#0f0e0c", borderRadius:8, padding:"8px 10px", fontWeight:700, cursor:"pointer" }}>
    Mes actual
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

            <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, marginBottom:14 }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-end", gap:10, flexWrap:"wrap", marginBottom:12 }}>
                <div>
                  <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent }}>🏦 Cartola bancaria</div>
                  <div style={{ color:COLORS.muted, fontSize:12 }}>Importa Excel/CSV del banco, revisa movimientos y conviértelos en abonos, pagos o asientos.</div>
                </div>
                <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"flex-end" }}>
                  <div>
                    <label style={{ display:"block", fontSize:10, color:COLORS.muted, marginBottom:5 }}>Mes cartola</label>
                    <input type="month" value={mesCartola} onChange={e=>setMesCartola(e.target.value)} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }} />
                  </div>
                  <button onClick={() => document.getElementById("importar-cartola-bancaria")?.click()} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>Importar cartola</button>
                </div>
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, marginBottom:12 }}>
                <StatCard label="Ingresos cartola" value={fmt(totalIngresosCartola)} icon="⬆️" color={COLORS.success}/>
                <StatCard label="Egresos cartola" value={fmt(totalEgresosCartola)} icon="⬇️" color={COLORS.danger}/>
                <StatCard label="Por revisar" value={cartolaPorRevisar.length} icon="🔎" color={cartolaPorRevisar.length ? COLORS.warning : COLORS.success}/>
              </div>

              <div style={{ display:"flex", justifyContent:"space-between", gap:10, flexWrap:"wrap", alignItems:"center", marginBottom:10 }}>
                <input value={busquedaCartola} onChange={e=>setBusquedaCartola(e.target.value)} placeholder="Buscar en cartola..." style={{ minWidth:220, flex:"0 1 320px", background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }} />
                <span style={{ color:COLORS.muted, fontSize:12 }}>{cartolaFiltrada.length} movimientos</span>
              </div>

              {cartolaFiltrada.length === 0 ? (
                <div style={{ color:COLORS.muted, fontSize:12 }}>Todavía no hay cartola importada para este mes.</div>
              ) : (
                <div style={{ overflowX:"auto", maxHeight:360, overflowY:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, minWidth:900 }}>
                    <thead>
                      <tr style={{ color:COLORS.muted, textAlign:"left" }}>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Fecha</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Descripción</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Cargo</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Abono</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Categoría</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Estado</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Acciones</th>
                      </tr>
                    </thead>
                    <tbody>
                      {cartolaFiltrada.slice(0, 80).map(m => (
                        <tr key={m.id}>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{fmtDate(m.fecha)}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, color:COLORS.text, fontWeight:700 }}>{m.descripcion}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.danger }}>{Number(m.cargo || 0) ? fmt(m.cargo) : "-"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.success }}>{Number(m.abono || 0) ? fmt(m.abono) : "-"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{m.categoria || "Por revisar"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, color:m.estado === "conciliado" ? COLORS.success : COLORS.warning }}>{m.estado === "conciliado" ? "Conciliado" : "Por revisar"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>
                            <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                              {Number(m.abono || 0) > 0 && <button onClick={() => registrarCartolaComoAbonoNV(m)} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11 }}>Abono NV</button>}
                              {Number(m.cargo || 0) > 0 && <button onClick={() => registrarCartolaComoPagoCuenta(m)} style={{ background:COLORS.warning, color:"#111", border:"none", borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11 }}>Pago CxP</button>}
                              <button onClick={() => registrarCartolaComoMovimientoContable(m)} style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:7, padding:"6px 8px", cursor:"pointer", fontSize:11 }}>Asiento</button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
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

      </div>


      {modalCuentaPagar && (
        <div onClick={() => setModalCuentaPagar(false)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.65)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:9999, overflowY:"auto", padding:"20px 10px" }}>
          <form onSubmit={guardarCuentaPorPagar} onClick={e=>e.stopPropagation()} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:20, width:"460px", maxWidth:"94vw", color:COLORS.text }}>
            <h2 style={{ marginTop:0, color:COLORS.accent }}>Nueva cuenta por pagar</h2>
            <label>Proveedor / institución<input name="nombre" required style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Tipo<select name="tipo" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}>
              <option value="proveedor">Proveedor</option><option value="imposiciones">Imposiciones</option><option value="remuneraciones">Remuneraciones</option><option value="arriendo">Arriendo</option><option value="servicio">Servicio</option><option value="cheque">Cheque</option><option value="deuda antigua">Deuda antigua</option><option value="otro">Otro</option>
            </select></label>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
              <label>Monto total<input name="monto_total" type="number" required style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
              <label>Monto cuota<input name="monto_cuota" type="number" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
              <label>Fecha vencimiento<input name="fecha_vencimiento" type="date" defaultValue={new Date().toISOString().split("T")[0]} style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
              <label>Cuota actual<input name="cuota_actual" type="number" defaultValue="1" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
              <label>Total cuotas<input name="cuotas_totales" type="number" defaultValue="1" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            </div>
            <label>Detalle<textarea name="detalle" style={{ width:"100%", minHeight:70, boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}><button type="button" onClick={() => setModalCuentaPagar(false)}>Cancelar</button><button type="submit" style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"9px 14px", fontWeight:700 }}>Guardar</button></div>
          </form>
        </div>
      )}

      {modalAbonoCxc && (
        <div onClick={() => setModalAbonoCxc(null)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.65)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:9999, overflowY:"auto", padding:"20px 10px" }}>
          <form
            onSubmit={async (e) => {
              e.preventDefault();
              const fd = new FormData(e.currentTarget);
              const monto = Number(String(fd.get("monto") || "0").replace(/[^0-9]/g, ""));
              const fecha = fd.get("fecha") || new Date().toISOString().split("T")[0];
              const medioPago = fd.get("medio_pago") || "transferencia";
              const tipoAbono = fd.get("tipo_abono") || "abono";
              const observacion = fd.get("observacion") || "Registrado desde CxC";

              const ok = await registrarAbonoDocumento({
                nota: modalAbonoCxc.nota,
                monto,
                fecha,
                medioPago,
                tipoAbono,
                observacion
              });

              if (ok) setModalAbonoCxc(null);
            }}
            onClick={e=>e.stopPropagation()}
            style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:20, width:"430px", maxWidth:"94vw", color:COLORS.text }}
          >
            <h2 style={{ marginTop:0, color:COLORS.accent }}>Registrar abono CxC</h2>
            <p style={{ marginTop:0, lineHeight:1.45 }}>
              <b>{modalAbonoCxc.tipo} #{modalAbonoCxc.numero}</b><br/>
              <span style={{ color:COLORS.text }}>{modalAbonoCxc.cliente}</span><br/>
              <span style={{ color:COLORS.muted }}>Total: {fmt(modalAbonoCxc.total)} · Abonado: {fmt(modalAbonoCxc.abonado)} · Saldo: {fmt(modalAbonoCxc.saldo)}</span>
            </p>

            <label style={{ display:"block", fontSize:12, color:COLORS.muted }}>Monto
              <input name="monto" type="number" required defaultValue={modalAbonoCxc.saldo} style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }} />
            </label>

            <label style={{ display:"block", fontSize:12, color:COLORS.muted }}>Tipo de abono
              <select name="tipo_abono" defaultValue="abono_saldo" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}>
                <option value="abono_inicial">Abono inicial</option>
                <option value="abono">Abono parcial</option>
                <option value="abono_retiro_parcial">Abono retiro parcial</option>
                <option value="abono_saldo">Abono saldo</option>
              </select>
            </label>

            <label style={{ display:"block", fontSize:12, color:COLORS.muted }}>Fecha
              <input name="fecha" type="date" defaultValue={new Date().toISOString().split("T")[0]} style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }} />
            </label>

            <label style={{ display:"block", fontSize:12, color:COLORS.muted }}>Medio de pago
              <select name="medio_pago" defaultValue="transferencia" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}>
                <option value="transferencia">Transferencia</option>
                <option value="efectivo">Efectivo</option>
                <option value="cheque">Cheque</option>
              </select>
            </label>

            <label style={{ display:"block", fontSize:12, color:COLORS.muted }}>Observación
              <input name="observacion" defaultValue="Registrado desde CxC" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 14px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }} />
            </label>

            <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
              <button type="button" onClick={() => setModalAbonoCxc(null)} style={{ background:COLORS.subtle, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 14px", cursor:"pointer" }}>Cancelar</button>
              <button type="submit" style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"9px 14px", fontWeight:700, cursor:"pointer" }}>Guardar abono</button>
            </div>
          </form>
        </div>
      )}

      {modalAbonoCuentaPagar && (
        <div onClick={() => setModalAbonoCuentaPagar(null)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.65)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:9999, overflowY:"auto", padding:"20px 10px" }}>
          <form onSubmit={guardarAbonoCuentaPagar} onClick={e=>e.stopPropagation()} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:20, width:"420px", maxWidth:"94vw", color:COLORS.text }}>
            <h2 style={{ marginTop:0, color:COLORS.accent }}>Abono a cuenta</h2>
            <p><b>{modalAbonoCuentaPagar.nombre}</b><br/><span style={{ color:COLORS.muted }}>Saldo: {fmt(saldoCuentaPagar(modalAbonoCuentaPagar))}</span></p>
            <label>Monto<input name="monto" type="number" required defaultValue={modalAbonoCuentaPagar.monto_cuota || saldoCuentaPagar(modalAbonoCuentaPagar)} style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Fecha<input name="fecha" type="date" defaultValue={new Date().toISOString().split("T")[0]} style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Medio pago<select name="medio_pago" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}><option value="transferencia">Transferencia</option><option value="efectivo">Efectivo</option><option value="cheque">Cheque</option></select></label>
            <label>Observación<input name="observacion" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}><button type="button" onClick={() => setModalAbonoCuentaPagar(null)}>Cancelar</button><button type="submit" style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"9px 14px", fontWeight:700 }}>Guardar abono</button></div>
          </form>
        </div>
      )}

      {modalTrabajador && (
        <div onClick={() => setModalTrabajador(false)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.65)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:9999, overflowY:"auto", padding:"20px 10px" }}>
          <form onSubmit={guardarTrabajador} onClick={e=>e.stopPropagation()} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:20, width:"440px", maxWidth:"94vw", color:COLORS.text }}>
            <h2 style={{ marginTop:0, color:COLORS.accent }}>Nuevo trabajador</h2>
            <label>Nombre<input name="nombre" required style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>RUT<input name="rut" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Cargo<input name="cargo" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Fecha ingreso<input name="fecha_ingreso" type="date" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Sueldo base<input name="sueldo_base" type="number" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}/></label>
            <label>Estado<select name="estado" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}><option value="activo">Activo</option><option value="finiquitado">Finiquitado</option></select></label>
            <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}><button type="button" onClick={() => setModalTrabajador(false)}>Cancelar</button><button type="submit" style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"9px 14px", fontWeight:700 }}>Guardar</button></div>
          </form>
        </div>
      )}

      {trabajadorArchivo && (
        <div onClick={() => setTrabajadorArchivo(null)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.65)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:9999, overflowY:"auto", padding:"20px 10px" }}>
          <form onSubmit={subirDocumentoTrabajador} onClick={e=>e.stopPropagation()} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:20, width:"420px", maxWidth:"94vw", color:COLORS.text }}>
            <h2 style={{ marginTop:0, color:COLORS.accent }}>Subir documento</h2>
            <p><b>{trabajadorArchivo.trabajador.nombre}</b><br/><span style={{ color:COLORS.muted }}>{trabajadorArchivo.file.name}</span></p>
            <label>Tipo documento<select name="tipo" style={{ width:"100%", boxSizing:"border-box", padding:10, margin:"6px 0 10px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}>
              <option value="Contrato">Contrato</option><option value="Anexo">Anexo</option><option value="Liquidación">Liquidación</option><option value="Cotizaciones">Cotizaciones</option><option value="Licencia médica">Licencia médica</option><option value="Vacaciones">Vacaciones</option><option value="Finiquito">Finiquito</option><option value="Otro">Otro</option>
            </select></label>
            <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}><button type="button" onClick={() => setTrabajadorArchivo(null)}>Cancelar</button><button type="submit" style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"9px 14px", fontWeight:700 }}>Subir</button></div>
          </form>
        </div>
      )}

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
{modalEditarCotizacion && (
  <EditarCotizacionCompletaModal
    cotizacion={modalEditarCotizacion}
    detalles={detallesCotizaciones.filter(d =>
      d.cotizacion_id === Number(String(modalEditarCotizacion.id).replace("supabase-", ""))
    )}
    onClose={() => setModalEditarCotizacion(null)}
    onSave={guardarEdicionCompletaCotizacion}
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
{modalFacturaXml && previewFacturaXml && (
  <FacturaXmlModal
    factura={previewFacturaXml}
    onClose={() => { setModalFacturaXml(false); setPreviewFacturaXml(null); }}
    onConfirm={confirmarImportacionFacturaXml}
  />
)}
{modalPreviewOC && (
  <PreviewOCModal
    productos={previewOC}
    documentos={previewOCDocumentos}
    onClose={() => { setModalPreviewOC(false); setPreviewOCDocumentos([]); }}
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

{modalEditarNV && (
  <EditarNVCompletaModal
    nota={modalEditarNV}
    detalles={detallesNotasVenta.filter(d => String(d.nota_venta_numero) === String(modalEditarNV.numero))}
    onClose={() => setModalEditarNV(null)}
    onSave={guardarEdicionCompletaNV}
  />
)}
    {modalNV && (
  <GestionNVModal
  nota={modalNV}
  abonos={abonosNV.filter(a => a.nota_venta_id === Number(String(modalNV.id).replace("supabase-", "")))}
  detalles={detallesNotasVenta.filter(d => String(d.nota_venta_numero) === String(modalNV.numero))}
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

      {modalClasificarCartola && (
        <div
          onClick={cerrarClasificacionCartola}
          style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.62)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:10000, padding:"20px 10px", overflowY:"auto" }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:14, padding:18, width:"min(920px, 96vw)", color:COLORS.text, boxSizing:"border-box" }}
          >
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:12, marginBottom:14 }}>
              <div>
                <h2 style={{ margin:"0 0 4px", color:COLORS.accent }}>Clasificar movimiento de cartola</h2>
                <div style={{ fontSize:12, color:COLORS.muted }}>{fmtDate(modalClasificarCartola.fecha)} · {modalClasificarCartola.descripcion}</div>
                <div style={{ marginTop:6, fontSize:18, fontWeight:900, color:Number(modalClasificarCartola.abono || 0) > 0 ? COLORS.success : COLORS.danger }}>
                  {Number(modalClasificarCartola.abono || 0) > 0 ? fmt(modalClasificarCartola.abono) : fmt(modalClasificarCartola.cargo)}
                </div>
              </div>
              <button onClick={cerrarClasificacionCartola} style={{ background:COLORS.danger, color:"#fff", border:"none", borderRadius:8, padding:"8px 10px", cursor:"pointer", fontWeight:800 }}>Cerrar</button>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"minmax(260px, 360px) 1fr", gap:14 }}>
              <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:12 }}>
                <div style={{ fontSize:12, color:COLORS.accent, fontWeight:800, marginBottom:8 }}>1. Categoría</div>
                <input
                  value={busquedaCategoriaCartola}
                  onChange={e=>setBusquedaCategoriaCartola(e.target.value)}
                  placeholder="Buscar categoría..."
                  style={{ width:"100%", boxSizing:"border-box", background:COLORS.card, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 10px", marginBottom:8 }}
                />
                <div style={{ display:"grid", gap:6, maxHeight:360, overflowY:"auto" }}>
                  {categoriasCartolaFiltradas.map(c => (
                    <button
                      key={c.nombre}
                      onClick={() => setCategoriaCartolaSeleccionada(c.nombre)}
                      style={{ textAlign:"left", background:categoriaCartolaSeleccionada === c.nombre ? COLORS.accent : COLORS.card, color:categoriaCartolaSeleccionada === c.nombre ? "#111" : COLORS.text, border:`1px solid ${categoriaCartolaSeleccionada === c.nombre ? COLORS.accent : COLORS.border}`, borderRadius:8, padding:"9px 10px", cursor:"pointer" }}
                    >
                      <div style={{ fontWeight:900 }}>{c.nombre}</div>
                      <div style={{ fontSize:10.5, opacity:.82 }}>{c.descripcion}</div>
                    </button>
                  ))}
                </div>
              </div>

              <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:12 }}>
                <div style={{ fontSize:12, color:COLORS.accent, fontWeight:800, marginBottom:8 }}>2. Detalle</div>

                {categoriaCartolaSeleccionada === "Abono CxC" ? (
                  <div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 170px", gap:10, marginBottom:10 }}>
                      <input
                        value={busquedaCxcCartola}
                        onChange={e=>setBusquedaCxcCartola(e.target.value)}
                        placeholder="Buscar NV, barrán o cliente..."
                        style={{ width:"100%", boxSizing:"border-box", background:COLORS.card, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 10px" }}
                      />
                      <select
                        value={tipoAbonoCartola}
                        onChange={e=>setTipoAbonoCartola(e.target.value)}
                        style={{ width:"100%", boxSizing:"border-box", background:COLORS.card, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"9px 10px" }}
                      >
                        <option value="abono_inicial">Abono inicial</option>
                        <option value="abono">Abono parcial</option>
                        <option value="abono_retiro_parcial">Abono retiro parcial</option>
                        <option value="abono_saldo">Abono saldo</option>
                      </select>
                    </div>

                    {cxcSeleccionadaCartola && (
                      <div style={{ background:"rgba(90,158,111,.12)", border:`1px solid ${COLORS.success}`, borderRadius:10, padding:10, marginBottom:10 }}>
                        <div style={{ color:COLORS.success, fontWeight:900 }}>Seleccionado: {cxcSeleccionadaCartola.tipo} #{cxcSeleccionadaCartola.numero}</div>
                        <div style={{ fontSize:12, color:COLORS.text }}>{cxcSeleccionadaCartola.cliente}</div>
                        <div style={{ fontSize:12, color:COLORS.muted }}>Saldo actual: {fmt(cxcSeleccionadaCartola.saldo)}</div>
                      </div>
                    )}

                    <div style={{ maxHeight:300, overflowY:"auto", display:"grid", gap:6 }}>
                      {opcionesCxcCartola.slice(0, 60).map(item => (
                        <button
                          key={item.id}
                          onClick={() => setCxcSeleccionadaCartola(item)}
                          style={{ textAlign:"left", background:cxcSeleccionadaCartola?.id === item.id ? COLORS.accent : COLORS.card, color:cxcSeleccionadaCartola?.id === item.id ? "#111" : COLORS.text, border:`1px solid ${cxcSeleccionadaCartola?.id === item.id ? COLORS.accent : COLORS.border}`, borderRadius:8, padding:"9px 10px", cursor:"pointer" }}
                        >
                          <div style={{ display:"flex", justifyContent:"space-between", gap:10 }}>
                            <b>{item.tipo} #{item.numero} · {item.cliente}</b>
                            <b>{fmt(item.saldo)}</b>
                          </div>
                          <div style={{ fontSize:10.5, opacity:.82 }}>{item.estado.icon} {item.estado.label} {item.detalleTexto ? `· ${item.detalleTexto}` : ""}</div>
                        </button>
                      ))}
                    </div>
                  </div>
                ) : (
                  <div style={{ color:COLORS.muted, fontSize:13, lineHeight:1.5 }}>
                    {categoriaCartolaSeleccionada
                      ? "Esta categoría quedará registrada en la cartola. Si corresponde a Factura compra, se mostrará como flujo, pero no se duplicará con el resumen tributario manual del programa externo."
                      : "Selecciona una categoría desde el listado."}
                  </div>
                )}

                <div style={{ display:"flex", justifyContent:"flex-end", gap:8, marginTop:14 }}>
                  <button onClick={cerrarClasificacionCartola} style={{ background:COLORS.subtle, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", cursor:"pointer", fontWeight:800 }}>Cancelar</button>
                  <button onClick={guardarClasificacionCartola} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"9px 12px", cursor:"pointer", fontWeight:900 }}>Guardar clasificación</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      <LoginAdminModal
        open={modalLogin}
        onClose={() => { setModalLogin(false); setLoginUser(""); setLoginPass(""); }}
        usuario={loginUser}
        setUsuario={setLoginUser}
        password={loginPass}
        setPassword={setLoginPass}
        onLogin={hacerLoginAdmin}
        cargando={cargandoLogin}
      />
      <NuevoErrorModal
        open={modalNuevoError}
        onClose={() => { setModalNuevoError(false); setFormError({ nv: "", tipo: "", descripcion: "", dias: 0, monto: 0 }); setBusquedaNV(""); }}
        tiposError={tiposError}
        formError={formError}
        setFormError={setFormError}
        busquedaNV={busquedaNV}
        setBusquedaNV={setBusquedaNV}
        nvsDisponibles={nvsDisponibles}
        onGuardar={guardarNuevoError}
      />
      {/* Modal de edición de error */}
      {errorEditando && (
        <div
          onClick={() => setErrorEditando(null)}
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(0,0,0,.65)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 9999,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: COLORS.card,
              border: `1px solid ${COLORS.border}`,
              borderRadius: 14,
              padding: 22,
              width: "min(400px, 90vw)",
              color: COLORS.text,
            }}
          >
            <h2 style={{ marginTop: 0, color: COLORS.accent, fontSize: 16 }}>
              ✏️ Editar Error #{errorEditando.id}
            </h2>
            <p style={{ fontSize: 13, color: COLORS.muted, marginBottom: 14 }}>
              <b>{errorEditando.nota_venta_numero}</b> — {errorEditando.cliente}
            </p>

            <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
              Días de Retrabaljo
            </label>
            <input
              type="number"
              min="0"
              value={formEditError.dias}
              onChange={(e) => setFormEditError((prev) => ({ ...prev, dias: Number(e.target.value) }))}
              style={{
                width: "100%",
                padding: "10px 12px",
                borderRadius: 8,
                border: `1px solid ${COLORS.border}`,
                background: COLORS.surface,
                color: COLORS.text,
                boxSizing: "border-box",
                marginBottom: 14,
                fontSize: 13,
              }}
            />

            <label style={{ display: "block", marginBottom: 6, fontSize: 12, color: COLORS.muted }}>
              Monto Material Perdido
            </label>
            <input
              type="number"
              min="0"
              value={formEditError.monto}
              onChange={(e) => setFormEditError((prev) => ({ ...prev, monto: Number(e.target.value) }))}
              style={{
                width: "100%",
                padding: "10px 12px",
                borderRadius: 8,
                border: `1px solid ${COLORS.border}`,
                background: COLORS.surface,
                color: COLORS.text,
                boxSizing: "border-box",
                marginBottom: 16,
                fontSize: 13,
              }}
            />

            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <button
                onClick={() => setErrorEditando(null)}
                style={{
                  background: COLORS.subtle,
                  border: `1px solid ${COLORS.border}`,
                  color: COLORS.text,
                  borderRadius: 8,
                  padding: "9px 16px",
                  fontWeight: 700,
                  cursor: "pointer",
                  fontSize: 13,
                }}
              >
                Cancelar
              </button>
              <button
                onClick={() => {
                  editarError(errorEditando.id, {
                    dias_retrabaljo: Number(formEditError.dias),
                    monto_material_perdido: Number(formEditError.monto),
                  });
                  setErrorEditando(null);
                }}
                style={{
                  background: COLORS.accent,
                  border: "none",
                  color: "#0f0e0c",
                  borderRadius: 8,
                  padding: "9px 16px",
                  fontWeight: 700,
                  cursor: "pointer",
                  fontSize: 13,
                }}
              >
                Guardar Cambios
              </button>
            </div>
          </div>
        </div>
      )}
      </>
      )}
    </div>
  );
}
