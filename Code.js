/**
 * ═══════════════════════════════════════════════════════════════
 * SCRIPT CCG v2.6
 *
 * CHANGELOG v2.6:
 * - Notificación "Nuevo pedido en QC" con mention a Mauro cuando entra un pedido sin ORIGEN
 * - Se envía una sola vez por pedido (Script Properties)
 *
 * CHANGELOG v2.5:
 * - Refactor: una sola función enviarWhatsApp(mensaje, mentionJID?)
 * - Helper obtenerDatosPedido(), mensajes más naturales
 *
 * CHANGELOG v2.4:
 * - Sistema de mentions para alertar a Mauro
 * - Notificación cuando pedido entra a CCG / listo para QC (PRODUCCION/MIXTO)
 *
 * CHANGELOG v2.3:
 * - Origen "MIXTO" (como Producción). Validaciones QC para PRODUCCION y MIXTO
 *
 * v2.2: Archivado automático, cálculo de tiempos
 * ═══════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ══════════════════════════════════════════════════════════════

const CONFIG = {
  HOJA_CCG: "CCG",
  HOJA_METRICAS: "Metricas_QC",
  HOJA_APROBADOS: "Aprobados",
  
  // WhatsApp
  MAURO_JID: "18099530116@s.whatsapp.net", // Número de Mauro
  
  // Columnas visibles CCG (15 columnas)
  COL: {
    PED_ID: 1,         // A
    CLIENTE: 2,        // B
    PRODUCTO: 3,       // C
    COLOR: 4,          // D
    CANTIDAD: 5,       // E
    UNIDAD: 6,         // F
    GLS_TOTALES: 7,    // G
    ORIGEN: 8,         // H
    GLS_REALES: 9,     // I
    VISCOSIDAD: 10,    // J
    PH: 11,            // K
    DENSIDAD: 12,      // L
    ESTADO_QC: 13,     // M
    FECHA: 14,         // N
    RESPONSABLE: 15    // O
  }
};

// ══════════════════════════════════════════════════════════════
// FUNCIÓN PRINCIPAL - onEdit (MEJORADA PARA COPY/PASTE)
// ══════════════════════════════════════════════════════════════

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const colStart = range.getColumn();
  const colEnd = colStart + range.getNumColumns() - 1; // Última columna editada
  
  // Filtro de seguridad
  if (sheet.getName() !== CONFIG.HOJA_CCG || row < 2) return;

  // Obtenemos el ID del pedido siempre para tenerlo disponible
  const pedId = sheet.getRange(row, CONFIG.COL.PED_ID).getValue();
  if (!pedId) return;

  // ═══════════════════════════════════════════════════════════
  // CASO 0: NUEVO PEDIDO EN CCG (ORIGEN vacío) → notificar una vez a Mauro
  // ═══════════════════════════════════════════════════════════
  const origen = sheet.getRange(row, CONFIG.COL.ORIGEN).getValue();
  if (!origen) {
    const key = "NUEVO_PED_" + String(pedId).replace(/\s/g, "_");
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty(key)) {
      notificarNuevoPedido(sheet, row);
      props.setProperty(key, "1");
    }
  }

  // ═══════════════════════════════════════════════════════════
  // CASO 1: DETECTAR CAMBIO EN "ORIGEN" (Incluso si es Copy/Paste)
  // ═══════════════════════════════════════════════════════════
  // Verificamos si la columna ORIGEN (8) cae dentro del rango editado
  if (CONFIG.COL.ORIGEN >= colStart && CONFIG.COL.ORIGEN <= colEnd) {
    // IMPORTANTE: Leemos el valor directo de la celda, no de e.value
    // porque e.value falla si pegaste un rango de celdas.
    const valorOrigen = sheet.getRange(row, CONFIG.COL.ORIGEN).getValue();
    
    // Solo ejecutamos si hay un valor real
    if (valorOrigen) {
      handleOrigenChange(sheet, row, valorOrigen, pedId);
    }
  }

  // ═══════════════════════════════════════════════════════════
  // CASO 2: Mauro intenta APROBAR (Columna 13)
  // ═══════════════════════════════════════════════════════════
  // Verificamos si la columna ESTADO_QC (13) cae dentro del rango editado
  if (CONFIG.COL.ESTADO_QC >= colStart && CONFIG.COL.ESTADO_QC <= colEnd) {
    const valorEstado = sheet.getRange(row, CONFIG.COL.ESTADO_QC).getValue();
    
    if (valorEstado === "APROBADO") {
      handleAprobacion(sheet, row, range, pedId);
    }
  }
}

// ══════════════════════════════════════════════════════════════
// HANDLERS
// ══════════════════════════════════════════════════════════════

function handleOrigenChange(sheet, row, origen, pedId) {
  // Validación ampliada para incluir MIXTO
  if (origen !== "STOCK" && origen !== "PRODUCCION" && origen !== "MIXTO") return;
  
  // Guardar timestamp en Metricas
  guardarTimestampOrigen(pedId);
  
  // ─────────────────────────────────────────────────────────
  // STOCK → Auto-aprobar
  // ─────────────────────────────────────────────────────────
  if (origen === "STOCK") {
    sheet.getRange(row, CONFIG.COL.ESTADO_QC).setValue("APROBADO");
    
    // Fecha automática
    const tsAprobado = new Date();
    sheet.getRange(row, CONFIG.COL.FECHA).setValue(tsAprobado);
    
    // Calcular y guardar métricas
    const metricas = calcularYGuardarMetricas(sheet, row, pedId);
    
    // Notificar (SIN mention - aprobación automática)
    notificarStockAprobado(sheet, row, metricas);
    
    Logger.log(`✅ ${pedId} - STOCK auto-aprobado`);
  }
  
  // ─────────────────────────────────────────────────────────
  // PRODUCCIÓN o MIXTO → Notificar pendiente CON MENTION
  // ─────────────────────────────────────────────────────────
  else if (origen === "PRODUCCION" || origen === "MIXTO") {
    notificarProduccionPendiente(sheet, row, origen);
    Logger.log(`🏭 ${pedId} - ${origen} pendiente de datos`);
  }
}

function handleAprobacion(sheet, row, range, pedId) {
  const origen = sheet.getRange(row, CONFIG.COL.ORIGEN).getValue();
  
  // ═══════════════════════════════════════════════════════════
  // VALIDACIÓN CRÍTICA: Origen debe estar definido
  // ═══════════════════════════════════════════════════════════
  if (origen !== "STOCK" && origen !== "PRODUCCION" && origen !== "MIXTO") {
    SpreadsheetApp.getUi().alert(
      "⛔ ERROR: ORIGEN NO DEFINIDO\n\n" +
      "No puede aprobar sin especificar el origen.\n\n" +
      "DEBE LLENAR PRIMERO:\n" +
      "• Columna H (Origen): STOCK, PRODUCCION o MIXTO\n\n" +
      "Después podrá aprobar."
    );
    range.setValue("PENDIENTE");
    Logger.log(`⛔ ${pedId} - Aprobación bloqueada: Origen = ${origen}`);
    return;
  }
  
  // Si es STOCK, ya se aprobó automáticamente
  if (origen === "STOCK") {
    Logger.log(`ℹ️ ${pedId} - STOCK ya fue auto-aprobado`);
    return;
  }
  
  // ─────────────────────────────────────────────────────────
  // PRODUCCIÓN o MIXTO → Validar campos obligatorios
  // ─────────────────────────────────────────────────────────
  if (origen === "PRODUCCION" || origen === "MIXTO") {
    const glsReales = sheet.getRange(row, CONFIG.COL.GLS_REALES).getValue();
    const viscosidad = sheet.getRange(row, CONFIG.COL.VISCOSIDAD).getValue();
    const pH = sheet.getRange(row, CONFIG.COL.PH).getValue();
    
    // Validación de campos técnicos
    if (!glsReales || !viscosidad || !pH) {
      SpreadsheetApp.getUi().alert(
        "⛔ FALTAN DATOS OBLIGATORIOS\n\n" +
        `Para origen ${origen} debe llenar:\n` +
        "• Gls_Reales (Col I)\n" +
        "• Viscosidad (Col J)\n" +
        "• pH (Col K)\n\n" +
        "Estado revertido a PENDIENTE."
      );
      range.setValue("PENDIENTE");
      Logger.log(`⛔ ${pedId} - Aprobación bloqueada: Faltan datos técnicos`);
      return;
    }
    
    // Todo OK → Aprobar
    const tsAprobado = new Date();
    sheet.getRange(row, CONFIG.COL.FECHA).setValue(tsAprobado);
    
    // Calcular y guardar métricas
    const metricas = calcularYGuardarMetricas(sheet, row, pedId);
    
    // Notificar (SIN mention - ya fue procesado)
    notificarProduccionAprobada(sheet, row, metricas, origen);
    
    Logger.log(`✅ ${pedId} - ${origen} aprobada`);
  }
}

// ══════════════════════════════════════════════════════════════
// MÉTRICAS - HOJA SEPARADA
// ══════════════════════════════════════════════════════════════

function guardarTimestampOrigen(pedId) {
  const metricasSheet = obtenerHojaMetricas();
  const data = metricasSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === pedId) {
      metricasSheet.getRange(i + 1, 7).setValue(new Date()); // G: TS_Origen_Llenado
      Logger.log(`📊 ${pedId} - TS_Origen_Llenado guardado`);
      return;
    }
  }
}

function calcularYGuardarMetricas(sheet, row, pedId) {
  const metricasSheet = obtenerHojaMetricas();
  const data = metricasSheet.getDataRange().getValues();
  
  let metricasRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === pedId) {
      metricasRow = i + 1;
      break;
    }
  }
  
  if (metricasRow === -1) {
    Logger.log(`⚠️ ${pedId} - No encontrado en Metricas_QC`);
    return null;
  }
  
  // Obtener timestamps
  const tsCreado = metricasSheet.getRange(metricasRow, 6).getValue();         // F
  const tsOrigenLlenado = metricasSheet.getRange(metricasRow, 7).getValue(); // G
  const tsAprobado = new Date();
  
  // Guardar TS_Aprobado
  metricasSheet.getRange(metricasRow, 8).setValue(tsAprobado); // H
  
  if (!tsCreado) {
    Logger.log(`⚠️ ${pedId} - TS_Creado no existe`);
    return null;
  }
  
  // Cálculos
  const tiempoTotal = (tsAprobado - tsCreado) / 1000 / 60;
  let tiempoCalidad = 0;
  let tiempoProduccion = tiempoTotal;
  
  if (tsOrigenLlenado) {
    tiempoCalidad = (tsAprobado - tsOrigenLlenado) / 1000 / 60;
    tiempoProduccion = tiempoTotal - tiempoCalidad;
  }
  
  const metricas = {
    tiempoProduccionFmt: formatearTiempo(tiempoProduccion),
    tiempoCalidadFmt: formatearTiempo(tiempoCalidad),
    tiempoTotalFmt: formatearTiempo(tiempoTotal)
  };
  
  // Guardar métricas en hoja
  metricasSheet.getRange(metricasRow, 9).setValue(metricas.tiempoProduccionFmt);   // I
  metricasSheet.getRange(metricasRow, 10).setValue(metricas.tiempoCalidadFmt);     // J
  metricasSheet.getRange(metricasRow, 11).setValue(metricas.tiempoTotalFmt);       // K
  
  // Datos de CCG
  const cliente = sheet.getRange(row, CONFIG.COL.CLIENTE).getValue();
  const producto = sheet.getRange(row, CONFIG.COL.PRODUCTO).getValue();
  const color = sheet.getRange(row, CONFIG.COL.COLOR).getValue();
  const origen = sheet.getRange(row, CONFIG.COL.ORIGEN).getValue();
  const glsReales = sheet.getRange(row, CONFIG.COL.GLS_REALES).getValue() || "";
  const viscosidad = sheet.getRange(row, CONFIG.COL.VISCOSIDAD).getValue() || "";
  const pH = sheet.getRange(row, CONFIG.COL.PH).getValue() || "";
  const densidad = sheet.getRange(row, CONFIG.COL.DENSIDAD).getValue() || "";
  const responsable = sheet.getRange(row, CONFIG.COL.RESPONSABLE).getValue() || "";
  
  metricasSheet.getRange(metricasRow, 2).setValue(cliente);       // B
  metricasSheet.getRange(metricasRow, 3).setValue(producto);      // C
  metricasSheet.getRange(metricasRow, 4).setValue(color);         // D
  metricasSheet.getRange(metricasRow, 5).setValue(origen);        // E
  metricasSheet.getRange(metricasRow, 12).setValue(glsReales);    // L
  metricasSheet.getRange(metricasRow, 13).setValue(viscosidad);   // M
  metricasSheet.getRange(metricasRow, 14).setValue(pH);           // N
  metricasSheet.getRange(metricasRow, 15).setValue(densidad);     // O
  metricasSheet.getRange(metricasRow, 16).setValue(responsable);  // P
  
  const fechaRegistro = Utilities.formatDate(tsAprobado, "GMT-4", "dd/MM/yyyy");
  const horaAprobado = Utilities.formatDate(tsAprobado, "GMT-4", "HH:mm");
  
  metricasSheet.getRange(metricasRow, 17).setValue(fechaRegistro); // Q
  metricasSheet.getRange(metricasRow, 18).setValue(horaAprobado);  // R
  
  Logger.log(`📊 Métricas guardadas: ${pedId} - ${metricas.tiempoTotalFmt}`);
  
  return metricas;
}

function obtenerHojaMetricas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let metricasSheet = ss.getSheetByName(CONFIG.HOJA_METRICAS);
  
  if (!metricasSheet) {
    metricasSheet = crearHojaMetricas(ss);
  }
  
  return metricasSheet;
}

function crearHojaMetricas(ss) {
  const metricasSheet = ss.insertSheet(CONFIG.HOJA_METRICAS);
  
  const headers = [
    "PED_ID", "Cliente", "Producto", "Color", "Origen",
    "TS_Creado", "TS_Origen_Llenado", "TS_Aprobado",
    "Tiempo_Produccion", "Tiempo_Calidad", "Tiempo_Total",
    "Gls_Reales", "Viscosidad", "pH", "Densidad", "Responsable",
    "Fecha_Registro", "Hora_Aprobado"
  ];
  
  metricasSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formato header
  metricasSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1E3A8A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // Ocultar hoja
  metricasSheet.hideSheet();
  
  Logger.log(`📊 Hoja ${CONFIG.HOJA_METRICAS} creada automáticamente`);
  
  return metricasSheet;
}

function formatearTiempo(minutos) {
  if (!minutos || minutos <= 0) return "";
  
  const horas = Math.floor(minutos / 60);
  const mins = Math.round(minutos % 60);
  
  if (horas > 0) {
    return `${horas}h ${mins}min`;
  }
  return `${mins}min`;
}

// ══════════════════════════════════════════════════════════════
// SISTEMA DE ARCHIVADO APROBADOS
// ══════════════════════════════════════════════════════════════

function moverAprobados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCCG = ss.getSheetByName(CONFIG.HOJA_CCG);
  const shAprobados = ss.getSheetByName(CONFIG.HOJA_APROBADOS);
  
  if (!shCCG || !shAprobados) {
    Logger.log("❌ Falta hoja CCG o Aprobados");
    return;
  }
  
  const data = shCCG.getDataRange().getValues();
  const filasAMover = [];
  const filasAEliminar = [];
  
  // BUSCAR APROBADOS
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const estadoQC = row[CONFIG.COL.ESTADO_QC - 1]; // M
    
    if (estadoQC === "APROBADO") {
      const pedId = row[CONFIG.COL.PED_ID - 1];
      const cliente = row[CONFIG.COL.CLIENTE - 1];
      const producto = row[CONFIG.COL.PRODUCTO - 1];
      const color = row[CONFIG.COL.COLOR - 1];
      const cantidad = row[CONFIG.COL.CANTIDAD - 1];
      const unidad = row[CONFIG.COL.UNIDAD - 1];
      const glsTotales = row[CONFIG.COL.GLS_TOTALES - 1];
      const origen = row[CONFIG.COL.ORIGEN - 1];
      const glsReales = row[CONFIG.COL.GLS_REALES - 1];
      const viscosidad = row[CONFIG.COL.VISCOSIDAD - 1];
      const pH = row[CONFIG.COL.PH - 1];
      const densidad = row[CONFIG.COL.DENSIDAD - 1];
      const responsable = row[CONFIG.COL.RESPONSABLE - 1];
      const fechaAprobacion = row[CONFIG.COL.FECHA - 1];
      
      // BUSCAR FECHA DE INGRESO EN METRICAS_QC
      let fechaIngreso = "";
      const shMetricas = ss.getSheetByName(CONFIG.HOJA_METRICAS);
      
      if (shMetricas) {
        const dataMetricas = shMetricas.getDataRange().getValues();
        for (let j = 1; j < dataMetricas.length; j++) {
          if (dataMetricas[j][0] === pedId) {
            fechaIngreso = dataMetricas[j][5]; // F: TS_Creado
            break;
          }
        }
      }
      
      if (!fechaIngreso) fechaIngreso = fechaAprobacion;
      
      // CALCULAR TIEMPO EN CCG
      let tiempoEnCCG = "";
      if (fechaIngreso && fechaAprobacion) {
        const diffMs = new Date(fechaAprobacion) - new Date(fechaIngreso);
        const minutos = Math.round(diffMs / (1000 * 60));
        tiempoEnCCG = formatearTiempo(minutos);
      }
      
      const horaAprobado = fechaAprobacion 
        ? Utilities.formatDate(new Date(fechaAprobacion), "GMT-4", "HH:mm")
        : "";
      
      const filaAprobado = [
        pedId, cliente, producto, color, cantidad, unidad, glsTotales,
        origen, glsReales, viscosidad, pH, densidad, responsable,
        fechaIngreso, fechaAprobacion, tiempoEnCCG, horaAprobado
      ];
      
      filasAMover.push(filaAprobado);
      filasAEliminar.push(i + 1);
    }
  }
  
  // MOVER A APROBADOS
  if (filasAMover.length > 0) {
    const lastRow = shAprobados.getLastRow();
    shAprobados.getRange(lastRow + 1, 1, filasAMover.length, 17)
      .setValues(filasAMover);
    filasAEliminar.reverse().forEach(rowNum => {
      shCCG.deleteRow(rowNum);
    });
  }
}

// ══════════════════════════════════════════════════════════════
// NOTIFICACIONES WHATSAPP
// ══════════════════════════════════════════════════════════════

/** Devuelve { pedId, cliente, producto, color } y opcionalmente viscosidad, pH */
function obtenerDatosPedido(sheet, row) {
  return {
    pedId: sheet.getRange(row, CONFIG.COL.PED_ID).getValue(),
    cliente: sheet.getRange(row, CONFIG.COL.CLIENTE).getValue(),
    producto: sheet.getRange(row, CONFIG.COL.PRODUCTO).getValue(),
    color: sheet.getRange(row, CONFIG.COL.COLOR).getValue(),
    viscosidad: sheet.getRange(row, CONFIG.COL.VISCOSIDAD).getValue(),
    pH: sheet.getRange(row, CONFIG.COL.PH).getValue()
  };
}

/** Respaldo: notificación cuando alguien EDITA en CCG una fila con ORIGEN vacío. La notificación principal "Nuevo pedido → QC" se envía desde QCBridge (CPG) al mandar a Calidad. */
function notificarNuevoPedido(sheet, row) {
  const d = obtenerDatosPedido(sheet, row);
  const cantidad = sheet.getRange(row, CONFIG.COL.CANTIDAD).getValue();
  const unidad = sheet.getRange(row, CONFIG.COL.UNIDAD).getValue();
  const cant = cantidad && unidad ? `${cantidad} ${unidad}` : cantidad || "—";
  const msg = `Nuevo pedido en QC: ${d.pedId} — ${d.cliente}, ${d.producto} ${d.color} (${cant}). Por favor indicar ORIGEN en CCG: ¿salió de STOCK o viene de PRODUCCIÓN/MIXTO?`;
  enviarWhatsApp(msg, CONFIG.MAURO_JID);
}

function notificarStockAprobado(sheet, row, metricas) {
  const d = obtenerDatosPedido(sheet, row);
  let msg = `📦 Listo: ${d.pedId} — ${d.cliente}, ${d.producto} ${d.color} (stock).`;
  if (metricas && metricas.tiempoTotalFmt) {
    msg += ` Tiempo: ${metricas.tiempoTotalFmt}.`;
  }
  msg += ` Listo para despachar.`;
  enviarWhatsApp(msg);
}

function notificarProduccionPendiente(sheet, row, origen) {
  const d = obtenerDatosPedido(sheet, row);
  const msg = `${d.pedId} — ${d.cliente}, ${d.producto} ${d.color}. Entró como ${origen}; cuando tengan los datos de QC (gls, viscosidad, pH) avisen para aprobar.`;
  enviarWhatsApp(msg, CONFIG.MAURO_JID);
}

function notificarProduccionAprobada(sheet, row, metricas, origen) {
  const d = obtenerDatosPedido(sheet, row);
  let msg = `✅ Aprobado: ${d.pedId} — ${d.cliente}, ${d.producto} ${d.color}. ${d.viscosidad} KU, pH ${d.pH}.`;
  if (metricas && metricas.tiempoTotalFmt) {
    msg += ` Tiempo total: ${metricas.tiempoTotalFmt}.`;
  }
  msg += ` Listo para despachar.`;
  enviarWhatsApp(msg);
}

// ══════════════════════════════════════════════════════════════
// ENVÍO WHATSAPP (mensaje; mentionJID opcional)
// ══════════════════════════════════════════════════════════════

function enviarWhatsApp(mensaje, mentionJID) {
  const props = PropertiesService.getScriptProperties();
  const WAS_TOKEN = props.getProperty('WAS_TOKEN');
  const GROUP_ID = props.getProperty('GROUP_GREQ_TECNICO');
  if (!WAS_TOKEN || !GROUP_ID) {
    Logger.log("⚠️ Token o Grupo no configurado");
    return;
  }
  const payload = { to: GROUP_ID, text: mensaje };
  if (mentionJID) payload.mentions = [mentionJID];
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${WAS_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const code = UrlFetchApp.fetch("https://www.wasenderapi.com/api/send-message", options).getResponseCode();
    Logger.log(`📱 WhatsApp enviado: ${code}`);
  } catch (error) {
    Logger.log(`❌ Error WhatsApp: ${error}`);
  }
}