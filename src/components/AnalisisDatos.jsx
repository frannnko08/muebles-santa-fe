import { useState, useEffect } from 'react';
import { supabase } from '../../lib/supabase';
import { LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

export function AnalisisDatos({ colors: COLORS }) {
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState(null);
  const [kpis, setKpis] = useState({ totalIngresos: 0, transacciones: 0, clientesUnicos: 0, promedio: 0 });
  const [porMes, setPorMes] = useState([]);
  const [todosClientes, setTodosClientes] = useState([]);
  const [tipos, setTipos] = useState([]);
  const [filtro, setFiltro] = useState('todo');
  const [topN, setTopN] = useState(5);
  const [clienteSeleccionado, setClienteSeleccionado] = useState(null);
  const [detalleCliente, setDetalleCliente] = useState([]);
  const [vistaActiva, setVistaActiva] = useState('dashboard');
  const [faltantes, setFaltantes] = useState({ cotizaciones: [], notas: [] });
  const [conversionPorMes, setConversionPorMes] = useState([]);
  const [clientesInactivos, setClientesInactivos] = useState([]);
  const [comparacionAnual, setComparacionAnual] = useState([]);
  const [proyeccion, setProyeccion] = useState([]);
  const [mesesInactivo, setMesesInactivo] = useState(6);

  useEffect(() => { cargarDatos(); }, [filtro]);

  async function cargarDatos() {
    setCargando(true);
    setError(null);
    try {
      const base = () => supabase
        .from('transacciones_historicas')
        .select('cliente, total, fecha, tipo, numero_cotizacion, numero_nota_venta')
        .eq('es_historico', true);

      const [res1, res2] = await Promise.all([
        base().range(0, 999),
        base().range(1000, 1999),
      ]);
      if (res1.error) throw res1.error;
      const data = [...(res1.data || []), ...(res2.data || [])];

      if (!data || data.length === 0) {
        setError('No se encontraron datos.');
        setCargando(false);
        return;
      }

      // Filtrar por fecha si corresponde
      const filtrado = filtro === 'todo' ? data : (() => {
        const desde = new Date();
        desde.setMonth(desde.getMonth() - parseInt(filtro));
        const desdeStr = desde.toISOString().split('T')[0];
        return data.filter(d => d.fecha >= desdeStr);
      })();

      // KPIs
      const ventas = filtrado.filter(d => d.tipo === 'nota_venta' || d.tipo === 'barran');
      const totalIngresos = ventas.reduce((s, d) => s + (Number(d.total) || 0), 0);
      const clientesUnicos = new Set(filtrado.map(d => d.cliente)).size;
      const promedio = ventas.length > 0 ? totalIngresos / ventas.length : 0;
      setKpis({ totalIngresos, transacciones: filtrado.length, clientesUnicos, promedio });

      // Por mes (ingresos)
      const mesMap = {};
      ventas.forEach(d => {
        const mes = (d.fecha || '').substring(0, 7);
        if (!mes) return;
        mesMap[mes] = (mesMap[mes] || 0) + (Number(d.total) || 0);
      });
      setPorMes(
        Object.entries(mesMap).sort(([a],[b]) => a.localeCompare(b)).map(([mes, total]) => {
          const [y, m] = mes.split('-');
          const nombres = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
          return { mes: `${nombres[parseInt(m)-1]} ${y.slice(-2)}`, mesISO: mes, total: Math.round(total) };
        })
      );

      // Top clientes
      const cliMap = {}, cliCount = {}, cliUltima = {};
      ventas.forEach(d => {
        cliMap[d.cliente] = (cliMap[d.cliente] || 0) + (Number(d.total) || 0);
        cliCount[d.cliente] = (cliCount[d.cliente] || 0) + 1;
        if (!cliUltima[d.cliente] || d.fecha > cliUltima[d.cliente]) cliUltima[d.cliente] = d.fecha;
      });
      const clientesOrdenados = Object.entries(cliMap).sort(([,a],[,b]) => b - a).map(([cliente, total], i) => ({
        rank: i + 1, cliente, total: Math.round(total),
        pct: totalIngresos > 0 ? ((total / totalIngresos) * 100).toFixed(1) : '0',
        transacciones: cliCount[cliente] || 0,
        ultimaCompra: cliUltima[cliente] || null
      }));
      setTodosClientes(clientesOrdenados);

      // Tipos
      const conteo = { nota_venta: 0, cotizacion: 0, barran: 0 };
      filtrado.forEach(d => { if (d.tipo in conteo) conteo[d.tipo]++; });
      setTipos([
        { name: 'Notas de Venta', value: conteo.nota_venta, color: '#5a9e6f' },
        { name: 'Cotizaciones', value: conteo.cotizacion, color: '#c8a45a' },
        { name: 'Barranes', value: conteo.barran, color: '#5a7ace' },
      ].filter(t => t.value > 0));

      // Faltantes
      const nvNums = [...new Set(filtrado.filter(d => d.numero_nota_venta && !isNaN(parseInt(d.numero_nota_venta))).map(d => parseInt(d.numero_nota_venta)))].sort((a,b) => a-b);
      const cotNums = [...new Set(filtrado.filter(d => d.numero_cotizacion && !isNaN(parseInt(d.numero_cotizacion))).map(d => parseInt(d.numero_cotizacion)))].sort((a,b) => a-b);
      setFaltantes({ notas: hallarFaltantes(nvNums), cotizaciones: hallarFaltantes(cotNums) });

      // === CONVERSIÓN POR MES ===
      const convMap = {};
      filtrado.forEach(d => {
        const mes = (d.fecha || '').substring(0, 7);
        if (!mes) return;
        if (!convMap[mes]) convMap[mes] = { cotizaciones: 0, ventas: 0 };
        if (d.tipo === 'cotizacion') convMap[mes].cotizaciones++;
        if (d.tipo === 'nota_venta' || d.tipo === 'barran') convMap[mes].ventas++;
      });
      setConversionPorMes(
        Object.entries(convMap).sort(([a],[b]) => a.localeCompare(b)).map(([mes, v]) => {
          const [y, m] = mes.split('-');
          const nombres = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
          const pct = v.cotizaciones > 0 ? Math.round((v.ventas / v.cotizaciones) * 100) : 0;
          return { mes: `${nombres[parseInt(m)-1]} ${y.slice(-2)}`, cotizaciones: v.cotizaciones, ventas: v.ventas, conversion: pct };
        })
      );

      // === COMPARACIÓN ANUAL ===
      const anioActual = new Date().getFullYear();
      const anioAnterior = anioActual - 1;
      const compMap = {};
      ventas.forEach(d => {
        if (!d.fecha) return;
        const anio = parseInt(d.fecha.substring(0, 4));
        const mes = parseInt(d.fecha.substring(5, 7));
        if (anio !== anioActual && anio !== anioAnterior) return;
        if (!compMap[mes]) compMap[mes] = { actual: 0, anterior: 0 };
        if (anio === anioActual) compMap[mes].actual += Number(d.total) || 0;
        if (anio === anioAnterior) compMap[mes].anterior += Number(d.total) || 0;
      });
      const nombres = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
      setComparacionAnual(
        Array.from({length: 12}, (_, i) => ({
          mes: nombres[i],
          [anioActual]: Math.round(compMap[i+1]?.actual || 0),
          [anioAnterior]: Math.round(compMap[i+1]?.anterior || 0),
        })).filter(d => d[anioActual] > 0 || d[anioAnterior] > 0)
      );

      // === PROYECCIÓN PRÓXIMOS 3 MESES ===
      const hoy = new Date();
      const proyArr = [];
      for (let i = 1; i <= 3; i++) {
        const fecha = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
        const mesNum = fecha.getMonth() + 1;
        const mismoMesHistorico = ventas.filter(d => d.fecha && parseInt(d.fecha.substring(5, 7)) === mesNum);
        const promMes = mismoMesHistorico.length > 0
          ? mismoMesHistorico.reduce((s, d) => s + (Number(d.total) || 0), 0) / new Set(mismoMesHistorico.map(d => d.fecha.substring(0, 4))).size
          : 0;
        proyArr.push({ mes: `${nombres[mesNum-1]} ${fecha.getFullYear().toString().slice(-2)}`, proyectado: Math.round(promMes) });
      }
      setProyeccion(proyArr);

      // === CLIENTES INACTIVOS (usando todos los datos sin filtro de fecha) ===
      const hoyStr = new Date().toISOString().split('T')[0];
      const inactivos = clientesOrdenados.filter(c => {
        if (!c.ultimaCompra) return false;
        const ultima = new Date(c.ultimaCompra);
        const diffMeses = (new Date() - ultima) / (1000 * 60 * 60 * 24 * 30);
        return diffMeses >= mesesInactivo;
      }).slice(0, 50);
      setClientesInactivos(inactivos);

    } catch (e) {
      setError('Error: ' + e.message);
    }
    setCargando(false);
  }

  function hallarFaltantes(nums) {
    if (nums.length < 2) return [];
    const min = nums[0], max = nums[nums.length - 1];
    if (max - min > 5000) return [];
    const set = new Set(nums), result = [];
    for (let i = min; i <= max && result.length < 200; i++) {
      if (!set.has(i)) result.push(i);
    }
    return result;
  }

  async function verDetalleCliente(cliente) {
    const { data } = await supabase
      .from('transacciones_historicas')
      .select('cliente, total, fecha, tipo, numero_cotizacion, numero_nota_venta')
      .eq('es_historico', true).eq('cliente', cliente).order('fecha', { ascending: false });
    setDetalleCliente(data || []);
    setClienteSeleccionado(cliente);
  }

  const fmt = n => '$' + Math.round(Number(n)).toLocaleString('es-CL');
  const fmtF = f => f ? new Date(f + 'T12:00:00').toLocaleDateString('es-CL') : '-';
  const topClientes = todosClientes.slice(0, topN);
  const anioActual = new Date().getFullYear();
  const anioAnterior = anioActual - 1;

  if (cargando) return (
    <div style={{ padding: 40, textAlign: 'center', color: COLORS.muted }}>
      <div style={{ fontSize: 32, marginBottom: 16 }}>⏳</div>
      <div>Cargando datos históricos...</div>
    </div>
  );

  if (error) return (
    <div style={{ padding: 40, color: COLORS.danger, fontSize: 14 }}>
      <b>❌ {error}</b><br/><br/>
      <button onClick={cargarDatos} style={{ background: COLORS.accent, color: '#111', border: 'none', borderRadius: 8, padding: '10px 20px', cursor: 'pointer', fontWeight: 700 }}>🔄 Reintentar</button>
    </div>
  );

  // DETALLE CLIENTE
  if (clienteSeleccionado) {
    const ventas = detalleCliente.filter(d => d.tipo !== 'cotizacion');
    const totalVentas = ventas.reduce((s, d) => s + (Number(d.total) || 0), 0);
    return (
      <div style={{ padding: 20, background: COLORS.bg, minHeight: '100vh' }}>
        <button onClick={() => setClienteSeleccionado(null)}
          style={{ background: 'transparent', border: `1px solid ${COLORS.border}`, color: COLORS.muted, borderRadius: 8, padding: '8px 16px', cursor: 'pointer', marginBottom: 20, fontSize: 13 }}>
          ← Volver
        </button>
        <h2 style={{ color: COLORS.accent, margin: '0 0 4px' }}>👤 {clienteSeleccionado}</h2>
        <p style={{ color: COLORS.muted, fontSize: 13, margin: '0 0 20px' }}>{detalleCliente.length} transacciones registradas</p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 14, marginBottom: 24 }}>
          {[
            { label: '💰 Total Ingresos', v: fmt(totalVentas) },
            { label: '✅ Notas / Barranes', v: ventas.length },
            { label: '📋 Cotizaciones', v: detalleCliente.filter(d => d.tipo === 'cotizacion').length },
            { label: '📊 Promedio por Venta', v: fmt(ventas.length > 0 ? totalVentas / ventas.length : 0) },
            { label: '📅 Primera compra', v: fmtF(detalleCliente[detalleCliente.length - 1]?.fecha) },
            { label: '📅 Última compra', v: fmtF(detalleCliente[0]?.fecha) },
          ].map(k => (
            <div key={k.label} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 14 }}>
              <div style={{ color: COLORS.muted, fontSize: 11, marginBottom: 6 }}>{k.label}</div>
              <div style={{ color: COLORS.accent, fontSize: 18, fontWeight: 900 }}>{k.v}</div>
            </div>
          ))}
        </div>
        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>📋 Historial</h3>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                  {['Fecha','Tipo','N° Cot.','N° NV','Total'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '8px', color: COLORS.muted, fontWeight: 700 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {detalleCliente.map((t, i) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                    <td style={{ padding: '10px 8px', color: COLORS.text }}>{fmtF(t.fecha)}</td>
                    <td style={{ padding: '10px 8px' }}>
                      <span style={{
                        background: t.tipo === 'nota_venta' ? '#1a3a25' : t.tipo === 'cotizacion' ? '#3a2e12' : '#12253a',
                        color: t.tipo === 'nota_venta' ? COLORS.success : t.tipo === 'cotizacion' ? COLORS.accent : '#5a9ace',
                        padding: '2px 8px', borderRadius: 6, fontSize: 11, fontWeight: 700
                      }}>
                        {t.tipo === 'nota_venta' ? 'NV' : t.tipo === 'cotizacion' ? 'COT' : 'BAR'}
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
        <button onClick={() => setVistaActiva('dashboard')}
          style={{ background: 'transparent', border: `1px solid ${COLORS.border}`, color: COLORS.muted, borderRadius: 8, padding: '8px 16px', cursor: 'pointer', marginBottom: 20, fontSize: 13 }}>
          ← Volver
        </button>
        <h2 style={{ color: COLORS.accent, margin: '0 0 4px' }}>🔍 Números Faltantes</h2>
        <p style={{ color: COLORS.muted, fontSize: 13, margin: '0 0 24px' }}>Correlativos no registrados en la base de datos.</p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: 20 }}>
          {[
            { titulo: '📋 Cotizaciones faltantes', nums: faltantes.cotizaciones, color: COLORS.accent },
            { titulo: '✅ Notas de Venta faltantes', nums: faltantes.notas, color: COLORS.success },
          ].map(({ titulo, nums, color }) => (
            <div key={titulo} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
              <h3 style={{ margin: '0 0 8px', color, fontSize: 14 }}>{titulo}</h3>
              <p style={{ color: COLORS.muted, fontSize: 12, margin: '0 0 12px' }}>
                {nums.length === 0 ? '✅ Sin faltantes detectados' : `${nums.length} faltante(s)`}
              </p>
              {nums.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, maxHeight: 300, overflowY: 'auto' }}>
                  {nums.map(n => (
                    <span key={n} style={{ background: COLORS.subtle, color: COLORS.text, padding: '3px 10px', borderRadius: 6, fontSize: 12, fontWeight: 700 }}>{n}</span>
                  ))}
                </div>
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
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
          <button onClick={() => setVistaActiva('faltantes')}
            style={{ background: COLORS.subtle, border: `1px solid ${COLORS.border}`, color: COLORS.text, borderRadius: 8, padding: '8px 14px', cursor: 'pointer', fontSize: 13, fontWeight: 700 }}>
            🔍 Ver Faltantes
          </button>
          <select value={filtro} onChange={e => setFiltro(e.target.value)}
            style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, color: COLORS.text, borderRadius: 8, padding: '8px 12px', fontSize: 13 }}>
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
          { label: '💰 Ingresos Totales (NV + Barranes)', v: fmt(kpis.totalIngresos) },
          { label: '📝 Total Transacciones', v: kpis.transacciones.toLocaleString('es-CL') },
          { label: '👥 Clientes Únicos', v: kpis.clientesUnicos.toLocaleString('es-CL') },
          { label: '📊 Promedio por Venta', v: fmt(kpis.promedio) },
        ].map(k => (
          <div key={k.label} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
            <div style={{ color: COLORS.muted, fontSize: 12, marginBottom: 8 }}>{k.label}</div>
            <div style={{ color: COLORS.accent, fontSize: 22, fontWeight: 900 }}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* Gráfico ingresos + comparación anual */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: 20, marginBottom: 24 }}>
        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>📈 Ingresos por Mes</h3>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={porMes}>
              <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
              <XAxis dataKey="mes" stroke={COLORS.muted} tick={{ fontSize: 11 }} />
              <YAxis stroke={COLORS.muted} tick={{ fontSize: 11 }} tickFormatter={v => '$' + (v/1000000).toFixed(1) + 'M'} />
              <Tooltip contentStyle={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, color: COLORS.text, fontSize: 12 }} formatter={v => [fmt(v), 'Ingresos']} />
              <Line type="monotone" dataKey="total" stroke={COLORS.accent} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>📅 {anioActual} vs {anioAnterior}</h3>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={comparacionAnual}>
              <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
              <XAxis dataKey="mes" stroke={COLORS.muted} tick={{ fontSize: 11 }} />
              <YAxis stroke={COLORS.muted} tick={{ fontSize: 11 }} tickFormatter={v => '$' + (v/1000000).toFixed(1) + 'M'} />
              <Tooltip contentStyle={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, color: COLORS.text, fontSize: 12 }} formatter={v => fmt(v)} />
              <Legend />
              <Bar dataKey={anioActual} fill={COLORS.accent} radius={[4,4,0,0]} />
              <Bar dataKey={anioAnterior} fill={COLORS.muted} radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Conversión + Tipos */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: 20, marginBottom: 24 }}>
        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 4px', color: COLORS.text, fontSize: 14 }}>🔄 Conversión Cotización → Venta por Mes</h3>
          <p style={{ color: COLORS.muted, fontSize: 12, margin: '0 0 12px' }}>Barras = cantidad | Línea = % conversión</p>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={conversionPorMes}>
              <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
              <XAxis dataKey="mes" stroke={COLORS.muted} tick={{ fontSize: 11 }} />
              <YAxis yAxisId="left" stroke={COLORS.muted} tick={{ fontSize: 11 }} />
              <YAxis yAxisId="right" orientation="right" stroke={COLORS.success} tick={{ fontSize: 11 }} tickFormatter={v => v + '%'} />
              <Tooltip contentStyle={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, color: COLORS.text, fontSize: 12 }} />
              <Legend />
              <Bar yAxisId="left" dataKey="cotizaciones" fill={COLORS.accentDim} name="Cotizaciones" radius={[4,4,0,0]} />
              <Bar yAxisId="left" dataKey="ventas" fill={COLORS.success} name="Ventas" radius={[4,4,0,0]} />
              <Line yAxisId="right" type="monotone" dataKey="conversion" stroke="#fff" strokeWidth={2} dot={false} name="% Conversión" />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 16px', color: COLORS.text, fontSize: 14 }}>🔄 Tipos de Transacciones</h3>
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={tipos} cx="50%" cy="50%" outerRadius={90} dataKey="value" label={({ name, value }) => `${name}: ${value}`} labelLine={false}>
                {tipos.map((e, i) => <Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip formatter={v => v.toLocaleString('es-CL')} />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Proyección + Clientes inactivos */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: 20, marginBottom: 24 }}>
        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <h3 style={{ margin: '0 0 4px', color: COLORS.text, fontSize: 14 }}>🔮 Proyección Próximos 3 Meses</h3>
          <p style={{ color: COLORS.muted, fontSize: 12, margin: '0 0 12px' }}>Basado en promedio histórico de cada mes</p>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12, marginBottom: 16 }}>
            {proyeccion.map(p => (
              <div key={p.mes} style={{ background: COLORS.subtle, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 14, textAlign: 'center' }}>
                <div style={{ color: COLORS.muted, fontSize: 12, marginBottom: 6 }}>{p.mes}</div>
                <div style={{ color: COLORS.accent, fontSize: 16, fontWeight: 900 }}>{fmt(p.proyectado)}</div>
              </div>
            ))}
          </div>
          <p style={{ color: COLORS.muted, fontSize: 11, margin: 0 }}>⚠️ Estimación basada en el historial. No considera factores externos.</p>
        </div>

        <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
            <h3 style={{ margin: 0, color: COLORS.text, fontSize: 14 }}>😴 Clientes Inactivos</h3>
            <select value={mesesInactivo} onChange={e => { setMesesInactivo(parseInt(e.target.value)); }}
              style={{ background: COLORS.subtle, border: `1px solid ${COLORS.border}`, color: COLORS.text, borderRadius: 8, padding: '4px 8px', fontSize: 12 }}>
              <option value={3}>+3 meses</option>
              <option value={6}>+6 meses</option>
              <option value={12}>+12 meses</option>
            </select>
          </div>
          <p style={{ color: COLORS.muted, fontSize: 12, margin: '4px 0 12px' }}>Clientes sin compras hace más de {mesesInactivo} meses</p>
          <div style={{ maxHeight: 220, overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                  {['Cliente', 'Última compra', 'Total histórico'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '6px 8px', color: COLORS.muted, fontWeight: 700 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {todosClientes.filter(c => {
                  if (!c.ultimaCompra) return false;
                  const diff = (new Date() - new Date(c.ultimaCompra + 'T12:00:00')) / (1000 * 60 * 60 * 24 * 30);
                  return diff >= mesesInactivo;
                }).slice(0, 20).map(c => (
                  <tr key={c.cliente} onClick={() => verDetalleCliente(c.cliente)}
                    style={{ borderBottom: `1px solid ${COLORS.border}`, cursor: 'pointer' }}
                    onMouseEnter={e => e.currentTarget.style.background = COLORS.subtle}
                    onMouseLeave={e => e.currentTarget.style.background = 'transparent'}>
                    <td style={{ padding: '8px', color: COLORS.text }}>{c.cliente} <span style={{ color: COLORS.muted, fontSize: 10 }}>→</span></td>
                    <td style={{ padding: '8px', color: COLORS.danger }}>{fmtF(c.ultimaCompra)}</td>
                    <td style={{ padding: '8px', color: COLORS.text, fontWeight: 700 }}>{fmt(c.total)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Top Clientes */}
      <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
          <h3 style={{ margin: 0, color: COLORS.text, fontSize: 14 }}>🏆 Top {topN} Clientes</h3>
          <select value={topN} onChange={e => setTopN(parseInt(e.target.value))}
            style={{ background: COLORS.subtle, border: `1px solid ${COLORS.border}`, color: COLORS.text, borderRadius: 8, padding: '6px 10px', fontSize: 12 }}>
            <option value={5}>Top 5</option>
            <option value={10}>Top 10</option>
            <option value={20}>Top 20</option>
            <option value={50}>Top 50</option>
          </select>
        </div>
        <p style={{ color: COLORS.muted, fontSize: 12, margin: '4px 0 16px' }}>Click en un cliente para ver su historial completo</p>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
              {['#','Cliente','Ventas','Total Ingresos','%'].map(h => (
                <th key={h} style={{ textAlign: h === 'Total Ingresos' || h === '%' ? 'right' : 'left', padding: '8px', color: COLORS.muted, fontWeight: 700 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {topClientes.map(c => (
              <tr key={c.rank} onClick={() => verDetalleCliente(c.cliente)}
                style={{ borderBottom: `1px solid ${COLORS.border}`, cursor: 'pointer' }}
                onMouseEnter={e => e.currentTarget.style.background = COLORS.subtle}
                onMouseLeave={e => e.currentTarget.style.background = 'transparent'}>
                <td style={{ padding: '12px 8px', color: COLORS.accent, fontWeight: 700 }}>#{c.rank}</td>
                <td style={{ padding: '12px 8px', color: COLORS.text, fontWeight: 600 }}>{c.cliente} <span style={{ color: COLORS.muted, fontSize: 11 }}>→</span></td>
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
