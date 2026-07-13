import { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';

export function AnalisisDatos({ colors: COLORS }) {
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState(null);
  const [kpis, setKpis] = useState({ totalIngresos: 0, transacciones: 0, clientesUnicos: 0, promedio: 0 });
  const [porMes, setPorMes] = useState([]);
  const [topClientes, setTopClientes] = useState([]);
  const [tipos, setTipos] = useState([]);
  const [filtro, setFiltro] = useState('todo');
  const [clienteSeleccionado, setClienteSeleccionado] = useState(null);
  const [detalleCliente, setDetalleCliente] = useState([]);
  const [faltantes, setFaltantes] = useState({ cotizaciones: [], notas: [], barranes: [] });
  const [vistaActiva, setVistaActiva] = useState('dashboard'); // 'dashboard' | 'faltantes'
  const [todosLosDatos, setTodosLosDatos] = useState([]);

  useEffect(() => { cargarDatos(); }, [filtro]);

  async function cargarDatos() {
    setCargando(true);
    setError(null);
    try {
      let query = supabase
        .from('transacciones_historicas')
        .select('id, cliente, total, fecha, tipo, numero_cotizacion, numero_nota_venta')
        .eq('es_historico', true);

      if (filtro !== 'todo') {
        const desde = new Date();
        desde.setMonth(desde.getMonth() - parseInt(filtro));
        query = query.gte('fecha', desde.toISOString().split('T')[0]);
      }

      const { data, error: err } = await query;
      if (err) throw err;
      if (!data || data.length === 0) {
        setError('No se encontraron datos.');
        setCargando(false);
        return;
      }

      setTodosLosDatos(data);

      // KPIs (excluir cotizaciones del ingreso)
      const ventas = data.filter(d => d.tipo === 'nota_venta' || d.tipo === 'barran');
      const totalIngresos = ventas.reduce((s, d) => s + (Number(d.total) || 0), 0);
      const clientesUnicos = new Set(data.map(d => d.cliente)).size;
      const promedio = ventas.length > 0 ? totalIngresos / ventas.length : 0;
      setKpis({ totalIngresos, transacciones: data.length, clientesUnicos, promedio });

      // Por mes (solo ventas reales)
      const mesMap = {};
      ventas.forEach(d => {
        const mes = (d.fecha || '').substring(0, 7);
        if (!mes) return;
        mesMap[mes] = (mesMap[mes] || 0) + (Number(d.total) || 0);
      });
      setPorMes(
        Object.entries(mesMap)
          .sort(([a], [b]) => a.localeCompare(b))
          .map(([mes, total]) => {
            const [y, m] = mes.split('-');
            const nombres = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
            return { mes: `${nombres[parseInt(m)-1]} ${y.slice(-2)}`, total: Math.round(total) };
          })
      );

      // Top clientes (por ingresos reales - solo notas y barranes)
      const cliMap = {};
      const cliCount = {};
      ventas.forEach(d => {
        cliMap[d.cliente] = (cliMap[d.cliente] || 0) + (Number(d.total) || 0);
        cliCount[d.cliente] = (cliCount[d.cliente] || 0) + 1;
      });
      setTopClientes(
        Object.entries(cliMap)
          .sort(([,a],[,b]) => b - a)
          .slice(0, 5)
          .map(([cliente, total], i) => ({
            rank: i + 1, cliente, total: Math.round(total),
            pct: ((total / totalIngresos) * 100).toFixed(1),
            transacciones: cliCount[cliente] || 0
          }))
      );

      // Tipos separados correctamente
      let nv = 0, cot = 0, bar = 0;
      data.forEach(d => {
        if (d.tipo === 'nota_venta') nv++;
        else if (d.tipo === 'cotizacion') cot++;
        else if (d.tipo === 'barran') bar++;
      });
      setTipos([
        { name: 'Notas de Venta', value: nv, color: '#5a9e6f' },
        { name: 'Cotizaciones', value: cot, color: '#c8a45a' },
        { name: 'Barranes', value: bar, color: '#5a7a9e' },
      ].filter(t => t.value > 0));

      // Detectar números faltantes
      calcularFaltantes(data);

    } catch (e) {
      setError('Error: ' + e.message);
    }
    setCargando(false);
  }

  function calcularFaltantes(data) {
    // Notas de venta: buscar en numero_nota_venta
    const numerosNV = data
      .filter(d => d.numero_nota_venta && d.numero_nota_venta !== '')
      .map(d => parseInt(d.numero_nota_venta))
      .filter(n => !isNaN(n))
      .sort((a, b) => a - b);

    // Cotizaciones: buscar en numero_cotizacion
    const numerosCot = data
      .filter(d => d.numero_cotizacion && d.numero_cotizacion !== '')
      .map(d => parseInt(d.numero_cotizacion))
      .filter(n => !isNaN(n))
      .sort((a, b) => a - b);

    const faltantesNV = encontrarFaltantes(numerosNV);
    const faltantesCot = encontrarFaltantes(numerosCot);

    setFaltantes({
      notas: faltantesNV,
      cotizaciones: faltantesCot,
      barranes: [] // Los barranes no tienen número correlativo claro
    });
  }

  function encontrarFaltantes(nums) {
    if (nums.length < 2) return [];
    const faltantes = [];
    for (let i = nums[0]; i <= nums[nums.length - 1]; i++) {
      if (!nums.includes(i)) faltantes.push(i);
    }
    // Limitar a 200 para no saturar
    return faltantes.slice(0, 200);
  }

  function verDetalleCliente(cliente) {
    const transacciones = todosLosDatos
      .filter(d => d.cliente === cliente)
      .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    setDetalleCliente(transacciones);
    setClienteSeleccionado(cliente);
  }

  const fmt = n => '$' + Math.round(Number(n)).toLocaleString('es-CL');
  const fmtFecha = f => f ? new Date(f + 'T12:00:00').toLocaleDateString('es-CL') : '-';

  if (cargando) return <div style={{ padding: 40, textAlign: 'center', color: COLORS.muted }}>⏳ Cargando datos históricos...</div>;
  if (error) return (
    <div style={{ padding: 40, color: COLORS.danger, fontSize: 14 }}>
      <b>❌ {error}</b><br/><br/>
      <button onClick={cargarDatos} style={{ background: COLORS.accent, color: '#111', border: 'none', borderRadius: 8, padding: '10px 20px', cursor: 'pointer', fontWeight: 700 }}>🔄 Reintentar</button>
    </div>
  );

  // MODAL DETALLE CLIENTE
  if (clienteSeleccionado) {
    const totalCliente = detalleCliente.reduce((s, d) => s + (Number(d.total) || 0), 0);
    const ventasCliente = detalleCliente.filter(d => d.tipo !== 'cotizacion');
    const totalVentas = ventasCliente.reduce((s, d) => s + (Number(d.total) || 0), 0);
    return (
      <div style={{ padding: 20, background: COLORS.bg, minHeight: '100vh' }}>
        <button
          onClick={() => setClienteSeleccionado(null)}
          style={{ background: 'transparent', border: `1px solid ${COLORS.border}`, color: COLORS.muted, borderRadius: 8, padding: '8px 16px', cursor: 'pointer', marginBottom: 20, fontSize: 13 }}
        >
          ← Volver al dashboard
        </button>
        <h2 style={{ color: COLORS.accent, margin: '0 0 4px' }}>👤 {clienteSeleccionado}</h2>
        <p style={{ color: COLORS.muted, fontSize: 13, margin: '0 0 20px' }}>{detalleCliente.length} transacciones registradas</p>

        {/* KPIs del cliente */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 16, marginBottom: 24 }}>
          {[
            { label: '💰 Total Ingresos', valor: fmt(totalVentas) },
            { label: '📝 Total Transacciones', valor: detalleCliente.length },
            { label: '✅ Notas / Barranes', valor: ventasCliente.length },
            { label: '📋 Cotizaciones', valor: detalleCliente.filter(d => d.tipo === 'cotizacion').length },
            { label: '📊 Promedio por Venta', valor: fmt(ventasCliente.length > 0 ? totalVentas / ventasCliente.length : 0) },
            { label: '📅 Primera compra', valor: fmtFecha(detalleCliente[detalleCliente.length - 1]?.fecha) },
            { label: '📅 Última compra', valor: fmtFecha(detalleCliente[0]?.fecha) },
          ].map(k => (
            <div key={k.label} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 14 }}>
              <div style={{ color: COLORS.muted, fontSize: 11, marginBottom: 6 }}>{k.label}</div>
              <div style={{ color: COLORS.accent, fontSize: 18, fontWeight: 900 }}>{k.valor}</div>
            </div>
          ))}
        </div>

        {/* Tabla de transacciones */}
        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>📋 Historial de Transacciones</h3>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                  {['Fecha', 'Tipo', 'N° Cotización', 'N° NV', 'Total'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '8px 8px', color: COLORS.muted, fontWeight: 700, whiteSpace: 'nowrap' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {detalleCliente.map((t, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                    <td style={{ padding: '10px 8px', color: COLORS.text }}>{fmtFecha(t.fecha)}</td>
                    <td style={{ padding: '10px 8px' }}>
                      <span style={{
                        background: t.tipo === 'nota_venta' ? '#1a3a25' : t.tipo === 'cotizacion' ? '#3a2e12' : '#12253a',
                        color: t.tipo === 'nota_venta' ? COLORS.success : t.tipo === 'cotizacion' ? COLORS.accent : '#5a9ace',
                        padding: '2px 8px', borderRadius: 6, fontSize: 11, fontWeight: 700
                      }}>
                        {t.tipo === 'nota_venta' ? 'Nota Venta' : t.tipo === 'cotizacion' ? 'Cotización' : 'Barran'}
                      </span>
                    </td>
                    <td style={{ padding: '10px 8px', color: COLORS.muted }}>{t.numero_cotizacion || '-'}</td>
                    <td style={{ padding: '10px 8px', color: COLORS.muted }}>{t.numero_nota_venta || '-'}</td>
                    <td style={{ padding: '10px 8px', color: COLORS.text, fontWeight: 700 }}>{fmt(t.total)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  // VISTA FALTANTES
  if (vistaActiva === 'faltantes') {
    return (
      <div style={{ padding: 20, background: COLORS.bg, minHeight: '100vh' }}>
        <button
          onClick={() => setVistaActiva('dashboard')}
          style={{ background: 'transparent', border: `1px solid ${COLORS.border}`, color: COLORS.muted, borderRadius: 8, padding: '8px 16px', cursor: 'pointer', marginBottom: 20, fontSize: 13 }}
        >
          ← Volver al dashboard
        </button>
        <h2 style={{ color: COLORS.accent, margin: '0 0 4px' }}>🔍 Números Faltantes</h2>
        <p style={{ color: COLORS.muted, fontSize: 13, margin: '0 0 24px' }}>
          Estos números correlativos no están en la base de datos. Puede que no se hayan extraído correctamente o no existan.
        </p>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: 20 }}>
          {[
            { titulo: '📋 Cotizaciones faltantes', nums: faltantes.cotizaciones, color: COLORS.accent },
            { titulo: '✅ Notas de Venta faltantes', nums: faltantes.notas, color: COLORS.success },
          ].map(({ titulo, nums, color }) => (
            <div key={titulo} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
              <h3 style={{ margin: '0 0 8px', color, fontSize: 14 }}>{titulo}</h3>
              <p style={{ color: COLORS.muted, fontSize: 12, margin: '0 0 12px' }}>
                {nums.length === 0 ? '✅ No se detectaron faltantes' : `${nums.length} número(s) faltante(s)`}
              </p>
              {nums.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, maxHeight: 300, overflowY: 'auto' }}>
                  {nums.map(n => (
                    <span key={n} style={{ background: COLORS.subtle, color: COLORS.text, padding: '3px 10px', borderRadius: 6, fontSize: 12, fontWeight: 700 }}>
                      {n}
                    </span>
                  ))}
                </div>
              )}
              {nums.length >= 200 && (
                <p style={{ color: COLORS.muted, fontSize: 11, marginTop: 8 }}>* Mostrando primeros 200 faltantes</p>
              )}
            </div>
          ))}
        </div>
      </div>
    );
  }

  // DASHBOARD PRINCIPAL
  return (
    <div style={{ padding: 20, background: COLORS.bg, minHeight: '100vh' }}>

      {/* Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20, flexWrap: 'wrap', gap: 10 }}>
        <h2 style={{ margin: 0, color: COLORS.accent, fontSize: 18 }}>📊 Análisis Histórico</h2>
        <div style={{ display: 'flex', gap: 10 }}>
          <button
            onClick={() => setVistaActiva('faltantes')}
            style={{ background: COLORS.subtle, border: `1px solid ${COLORS.border}`, color: COLORS.text, borderRadius: 8, padding: '8px 14px', cursor: 'pointer', fontSize: 13, fontWeight: 700 }}
          >
            🔍 Ver Faltantes ({faltantes.cotizaciones.length + faltantes.notas.length})
          </button>
          <select
            value={filtro}
            onChange={e => setFiltro(e.target.value)}
            style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, color: COLORS.text, borderRadius: 8, padding: '8px 12px', fontSize: 13 }}
          >
            <option value="todo">Todo el período</option>
            <option value="3">Últimos 3 meses</option>
            <option value="6">Últimos 6 meses</option>
            <option value="12">Últimos 12 meses</option>
            <option value="24">Últimos 24 meses</option>
          </select>
        </div>
      </div>

      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16, marginBottom: 24 }}>
        {[
          { label: '💰 Ingresos Totales (NV + Barranes)', valor: fmt(kpis.totalIngresos) },
          { label: '📝 Total Transacciones', valor: kpis.transacciones.toLocaleString('es-CL') },
          { label: '👥 Clientes Únicos', valor: kpis.clientesUnicos.toLocaleString('es-CL') },
          { label: '📊 Promedio por Venta', valor: fmt(kpis.promedio) },
        ].map(k => (
          <div key={k.label} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
            <div style={{ color: COLORS.muted, fontSize: 12, marginBottom: 8 }}>{k.label}</div>
            <div style={{ color: COLORS.accent, fontSize: 22, fontWeight: 900 }}>{k.valor}</div>
          </div>
        ))}
      </div>

      {/* Gráficos */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: 20, marginBottom: 24 }}>

        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>📈 Ingresos por Mes (NV + Barranes)</h3>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={porMes}>
              <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
              <XAxis dataKey="mes" stroke={COLORS.muted} tick={{ fontSize: 11 }} />
              <YAxis stroke={COLORS.muted} tick={{ fontSize: 11 }} tickFormatter={v => '$' + (v/1000000).toFixed(1) + 'M'} />
              <Tooltip
                contentStyle={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, color: COLORS.text, fontSize: 12 }}
                formatter={v => [fmt(v), 'Ingresos']}
              />
              <Line type="monotone" dataKey="total" stroke={COLORS.accent} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>🔄 Tipo de Transacciones</h3>
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={tipos} cx="50%" cy="50%" outerRadius={90} dataKey="value"
                label={({ name, value }) => `${name}: ${value}`} labelLine={false}>
                {tipos.map((e, i) => <Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip formatter={v => v.toLocaleString('es-CL')} />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Top 5 Clientes */}
      <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
        <h3 style={{ margin: '0 0 4px', color: COLORS.text, fontSize: 14 }}>🏆 Top 5 Clientes</h3>
        <p style={{ color: COLORS.muted, fontSize: 12, margin: '0 0 16px' }}>Haz click en un cliente para ver el detalle completo</p>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
              {['#', 'Cliente', 'Transacciones', 'Total Ingresos', '%'].map(h => (
                <th key={h} style={{ textAlign: h === 'Total Ingresos' || h === '%' ? 'right' : 'left', padding: '8px 8px', color: COLORS.muted, fontWeight: 700 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {topClientes.map(c => (
              <tr
                key={c.rank}
                onClick={() => verDetalleCliente(c.cliente)}
                style={{ borderBottom: `1px solid ${COLORS.border}`, cursor: 'pointer', transition: 'background 0.15s' }}
                onMouseEnter={e => e.currentTarget.style.background = COLORS.subtle}
                onMouseLeave={e => e.currentTarget.style.background = 'transparent'}
              >
                <td style={{ padding: '12px 8px', color: COLORS.accent, fontWeight: 700 }}>#{c.rank}</td>
                <td style={{ padding: '12px 8px', color: COLORS.text, fontWeight: 600 }}>
                  {c.cliente} <span style={{ color: COLORS.muted, fontSize: 11 }}>→ ver detalle</span>
                </td>
                <td style={{ padding: '12px 8px', color: COLORS.muted }}>{c.transacciones}</td>
                <td style={{ padding: '12px 8px', textAlign: 'right', color: COLORS.text, fontWeight: 700 }}>{fmt(c.total)}</td>
                <td style={{ padding: '12px 8px', textAlign: 'right', color: COLORS.success, fontWeight: 700 }}>{c.pct}%</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
