import { supabase } from './supabase'

export async function obtenerCotizaciones() {
  const { data, error } = await supabase
    .from('cotizaciones')
    .select('*')
    .order('numero', { ascending: false })

  if (error) {
    console.error('Error obteniendo cotizaciones:', error)
    return []
  }

  return data
}

export async function crearCotizacion(cliente, total) {
  const { data: ultimaCotizacion } = await supabase
    .from('cotizaciones')
    .select('numero')
    .order('numero', { ascending: false })
    .limit(1)

  const nuevoNumero = ultimaCotizacion?.[0]?.numero + 1 || 1

  const { data, error } = await supabase
    .from('cotizaciones')
    .insert([
      {
        numero: nuevoNumero,
        cliente: cliente,
        total: total,
        estado: 'pendiente'
      }
    ])
    .select()

  if (error) {
    console.error('Error creando cotización:', error)
    return null
  }

  console.log('Cotización creada:', data)
  return data
}
export async function aceptarCotizacion(cotizacionId) {
    const { data: cotizacionActual, error: errorConsulta } = await supabase
    .from('cotizaciones')
    .select('estado, nota_venta')
    .eq('id', cotizacionId)
    .single()

  if (errorConsulta) {
    console.error('Error consultando cotización:', errorConsulta)
    return null
  }

  if (cotizacionActual.estado === 'aceptada' || cotizacionActual.nota_venta) {
    alert('Esta cotización ya fue aceptada anteriormente')
    return null
  }
  const { data: ultimaNota } = await supabase
    .from('notas_venta')
    .select('numero')
    .order('numero', { ascending: false })
    .limit(1)

  const nuevoNumeroNV = ultimaNota?.[0]?.numero + 1 || 1

  const { error: errorNota } = await supabase
    .from('notas_venta')
    .insert([
      {
        numero: nuevoNumeroNV,
        cotizacion_id: cotizacionId
      }
    ])

  if (errorNota) {
    console.error('Error creando nota de venta:', errorNota)
    return null
  }

  const { error: errorCotizacion } = await supabase
    .from('cotizaciones')
    .update({
      estado: 'aceptada',
      fecha_aceptacion: new Date(),
      nota_venta: nuevoNumeroNV
    })
    .eq('id', cotizacionId)

  if (errorCotizacion) {
    console.error('Error actualizando cotización:', errorCotizacion)
    return null
  }

  console.log('Cotización aceptada y NV creada:', nuevoNumeroNV)

  return nuevoNumeroNV
}
export async function obtenerNotasVenta() {
  const { data, error } = await supabase
    .from('notas_venta')
    .select(`
  id,
  numero,
  fecha,
  cotizacion_id,
  cliente,
  total,
  cotizaciones (
    numero,
    cliente,
    total
  )
`)
    .order('numero', { ascending: false })

  if (error) {
    console.error('Error obteniendo notas de venta:', error)
    return []
  }

  return data
}
export async function editarNumeroNotaVenta(cotizacionId, nuevoNumeroNV) {
  const { error: errorCotizacion } = await supabase
    .from('cotizaciones')
    .update({
      nota_venta: nuevoNumeroNV
    })
    .eq('id', cotizacionId)

  if (errorCotizacion) {
    console.error('Error actualizando NV en cotización:', errorCotizacion)
    return false
  }

  const { error: errorNota } = await supabase
    .from('notas_venta')
    .update({
      numero: nuevoNumeroNV
    })
    .eq('cotizacion_id', cotizacionId)

  if (errorNota) {
    console.error('Error actualizando número de nota:', errorNota)
    return false
  }

  return true
}
export async function migrarNotasVenta(notas, cotizaciones) {
  for (const n of notas) {
    const { data: existeNV } = await supabase
      .from('notas_venta')
      .select('id')
      .eq('numero', n.numero)
      .limit(1)

    if (existeNV && existeNV.length > 0) {
      console.log(`NV ${n.numero} ya existe`);
      continue;
    }

    let cotizacionId = null;

    if (n.cotizacion) {
      const cotizacionRelacionada = cotizaciones.find(
        c => String(c.numero) === String(n.cotizacion)
      )

      if (cotizacionRelacionada) {
        const numeroCotizacion = Number(
          String(cotizacionRelacionada.numero).replace(/[^\d]/g, "")
        )

        const { data: cotizacionDB } = await supabase
          .from('cotizaciones')
          .select('id')
          .eq('numero', numeroCotizacion)
          .single()

        if (cotizacionDB) {
          cotizacionId = cotizacionDB.id;
        }
      }
    }

    const { error } = await supabase
      .from('notas_venta')
      .insert([
        {
          numero: Number(n.numero),
          cotizacion_id: cotizacionId,
          cliente: n.cliente,
          total: n.total,
          fecha: n.fecha
        }
      ])

    if (error) {
      console.error('Error migrando NV:', n.numero, error)
    } else {
      console.log('NV migrada:', n.numero)
    }
  }

  alert('Migración de notas terminada')
}
export async function importarNotaVentaExcel({ cotizacion, notaVenta, cliente, fecha, total }) {
  const numeroCotizacion = Number(String(cotizacion).replace(/[^\d]/g, ""));
  const numeroNotaVenta = Number(String(notaVenta).replace(/[^\d]/g, ""));
  const totalLimpio = Math.round(Number(total));

  let fechaFormateada = fecha;

  if (typeof fecha === "number") {
    const excelDate = new Date((fecha - 25569) * 86400 * 1000);
    fechaFormateada = excelDate.toISOString().split("T")[0];
  }

  const { data: notaExistente } = await supabase
    .from("notas_venta")
    .select("id")
    .eq("numero", numeroNotaVenta)
    .limit(1);

  if (notaExistente && notaExistente.length > 0) {
    console.log("Actualizando NV existente:", {
  numeroNotaVenta,
  cliente,
  fechaFormateada,
  totalLimpio
});
  const { error: errorUpdate } = await supabase
    .from("notas_venta")
    .update({
      cliente,
      fecha: fechaFormateada,
      total: totalLimpio,
    })
    .eq("numero", numeroNotaVenta);

  if (errorUpdate) {
    console.error("Error actualizando nota existente:", errorUpdate);
    alert("Error actualizando nota existente");
    return false;
  }

  alert("Nota de venta ya existía, se actualizó correctamente");
  return true;
}

  let cotizacionId = null;

  if (numeroCotizacion) {
    const { data: cotizacionDB } = await supabase
      .from("cotizaciones")
      .select("id")
      .eq("numero", numeroCotizacion)
      .limit(1);

    if (cotizacionDB && cotizacionDB.length > 0) {
      cotizacionId = cotizacionDB[0].id;
    }
  }

  const { error } = await supabase
    .from("notas_venta")
    .insert([
      {
        numero: numeroNotaVenta,
        cotizacion_id: cotizacionId,
        cliente,
        fecha: fechaFormateada,
        total: totalLimpio,
      },
    ]);

  if (error) {
    console.error("Error importando nota de venta:", error);
    alert("Error importando nota de venta");
    return false;
  }

  if (cotizacionId) {
    await supabase
      .from("cotizaciones")
      .update({
        estado: "aceptada",
        nota_venta: numeroNotaVenta,
        fecha_aceptacion: fechaFormateada,
      })
      .eq("id", cotizacionId);
  }

  alert("Nota de venta importada correctamente");
  return true;
}
export async function eliminarNotaVenta(notaVentaId) {
  const idReal = String(notaVentaId).replace("supabase-", "");

  const { error } = await supabase
    .from("notas_venta")
    .delete()
    .eq("id", idReal);

  if (error) {
    console.error("Error eliminando nota de venta:", error);
    alert("Error eliminando nota de venta");
    return false;
  }

  alert("Nota de venta eliminada");
  return true;
}
export async function importarCotizacionExcel({
  numero,
  cliente,
  fecha,
  total,
  detalles = []
}) {
  const numeroCotizacion = Number(String(numero).replace(/[^\d]/g, ""));
  const totalLimpio = Math.round(Number(total));

  let fechaFormateada = fecha;

  if (typeof fecha === "number") {
    const excelDate = new Date((fecha - 25569) * 86400 * 1000);
    fechaFormateada = excelDate.toISOString().split("T")[0];
  }

  const { data: existente } = await supabase
    .from("cotizaciones")
    .select("id")
    .eq("numero", numeroCotizacion)
    .limit(1);

  if (existente && existente.length > 0) {
    alert("Esta cotización ya existe");
    return false;
  }

  const { data: cotizacionCreada, error } = await supabase
  .from("cotizaciones")
  .insert([
    {
      numero: numeroCotizacion,
      cliente,
      fecha_creacion: fechaFormateada,
      total: totalLimpio,
      estado: "pendiente"
    }
  ])
  .select();

  if (error) {
    console.error("Error importando cotización:", error);
    alert("Error importando cotización");
    return false;
  }
  if (
  cotizacionCreada &&
  cotizacionCreada[0] &&
  detalles.length > 0
) {
  const detallesInsert = detalles.map((d, index) => ({
    cotizacion_id: cotizacionCreada[0].id,
    unidad: d.unidad,
    tipo: d.tipo,
    largo: d.largo,
    ancho: d.ancho,
    color: d.color,
    valor: d.valor,
    total: d.total,
    orden: index
  }));

  const { error: errorDetalles } = await supabase
    .from("cotizacion_detalles")
    .insert(detallesInsert);

  if (errorDetalles) {
    console.error("ERROR DETALLES:", errorDetalles);
    alert("La cotización se guardó, pero hubo error guardando detalles.");
    }
  }

  alert("Cotización importada correctamente");
  return true;
}
export async function obtenerDetallesCotizaciones() {
  const { data, error } = await supabase
    .from("cotizacion_detalles")
    .select("*")
    .order("orden", { ascending: true });

  if (error) {
    console.error("Error obteniendo detalles:", error);
    return [];
  }

  return data || [];
}