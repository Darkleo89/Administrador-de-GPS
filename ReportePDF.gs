// ============================================================
// SISTEMA DE GESTIÓN DE FLOTAS — ReportePDF.gs
// Generación de HTML maquetado → PDF → Drive
// ============================================================
/**
 * Genera un PDF y lo envía por correo a finanzas
 * @param {string} token - Token de sesión
 * @param {string} folio - Folio del registro
 * @returns {Object} { ok, pdfUrl, correoEnviado, error }
 */
function generarYEnviarPDF(token, folio) {
  try {
    // Usar la función base
    const resultado = _generarPDFBase(token, folio, {
      subirADrive: true,           // Subir a Drive
      actualizarBitacora: true,    // Actualizar URL en bitácora
      registrarAuditoria: true     // Registrar auditoría
    });

    if (!resultado.success) {
      return { 
        ok: false, 
        error: resultado.error 
      };
    }

    const { registro, pdfBlob, pdfUrl } = resultado;

    // ============================================================
    // ENVIAR CORREO A FINANZAS
    // ============================================================
    let correoEnviado = false;
    let errorCorreo = null;

    try {
      // Verificar que la función de correo exista
      if (typeof _enviarCorreoFinanzas === 'function') {
        const correoResultado = _enviarCorreoFinanzas(registro, pdfBlob, pdfUrl);
        correoEnviado = correoResultado?.success || false;
        if (!correoEnviado) {
          errorCorreo = correoResultado?.error || 'Error al enviar correo';
        }
      } else {
        console.warn('⚠️ _enviarCorreoFinanzas no está definida');
        errorCorreo = 'Función de correo no disponible';
      }
    } catch (err) {
      console.error('❌ Error al enviar correo:', err);
      errorCorreo = err.message;
    }

    // ============================================================
    // ACTUALIZAR ESTADO A "PAGADO"
    // ============================================================
    try {
      _actualizarEstadoRegistro(folio, 'Pagado');
      console.log('✅ Registro ' + folio + ' actualizado a Pagado');
    } catch (err) {
      console.warn('⚠️ No se pudo actualizar estado:', err.message);
      // No fallamos el proceso
    }

    // ============================================================
    // RETORNAR RESULTADO
    // ============================================================
    return {
      ok: true,
      folio: folio,
      pdfUrl: pdfUrl,
      correoEnviado: correoEnviado,
      errorCorreo: errorCorreo,
      mensaje: correoEnviado 
        ? 'PDF generado y correo enviado exitosamente' 
        : 'PDF generado pero no se pudo enviar el correo'
    };

  } catch (err) {
    console.error('❌ Error en generarYEnviarPDF:', err);
    return { 
      ok: false, 
      error: 'Error al generar y enviar PDF: ' + err.message 
    };
  }
}
/**
 * Actualiza el estado de un registro en la bitácora
 * @param {string} folio - Folio del registro
 * @param {string} nuevoEstado - Nuevo estado
 */
function _actualizarEstadoRegistro(folio, nuevoEstado) {
  try {
    const sheet = SHEETS.BITACORA();
    if (!sheet) {
      throw new Error('No se encontró la hoja de Bitácora');
    }

    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];
    const idxFolio = headers.indexOf('FOLIO');
    const idxEstado = headers.indexOf('ESTADO');

    if (idxFolio === -1 || idxEstado === -1) {
      throw new Error('Columnas FOLIO o ESTADO no encontradas');
    }

    for (var i = 1; i < datos.length; i++) {
      if ((datos[i][idxFolio] || '').toString() === folio) {
        sheet.getRange(i + 1, idxEstado + 1).setValue(nuevoEstado);
        console.log('✅ Estado actualizado a "' + nuevoEstado + '" para folio ' + folio);
        return true;
      }
    }

    throw new Error('No se encontró el folio ' + folio);

  } catch (err) {
    console.error('❌ Error al actualizar estado:', err);
    throw err;
  }
}

/**
 * Actualiza la URL del PDF en Bitácora
 * @param {string} folio - Folio del registro
 * @param {string} pdfUrl - URL del PDF
 */
function _actualizarURLPDF(folio, pdfUrl) {
  try {
    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      console.error('❌ No se encontró la hoja de Bitácora');
      return;
    }

    var datos = sheet.getDataRange().getValues();
    var headers = datos[0];
    var idxFolio = headers.indexOf('FOLIO');
    var idxPDFUrl = headers.indexOf('URL_PDF');

    if (idxFolio === -1 || idxPDFUrl === -1) {
      console.error('❌ No se encontraron las columnas necesarias');
      return;
    }

    var folioBusqueda = folio.toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var folioActual = (datos[i][idxFolio] || '').toString().toUpperCase().trim();
      if (folioActual === folioBusqueda) {
        sheet.getRange(i + 1, idxPDFUrl + 1).setValue(pdfUrl);
        console.log('✅ URL actualizada en Bitácora para:', folio);
        return;
      }
    }

  } catch (err) {
    console.error('❌ Error en _actualizarURLPDF:', err);
  }
}


/**
 * Obtiene un registro completo por folio mapeando dinámicamente sus columnas
 * @param {string} folio - Folio a buscar
 * @returns {Object|null} - Registro estructurado o null si no existe
 */
function _obtenerRegistroPorFolio(folio) {
  try {
    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      console.error('❌ No se encontró la hoja de Bitácora');
      return null;
    }

    var datos = sheet.getDataRange().getValues();
    var headers = datos[0];
    var idxFolio = headers.indexOf('FOLIO');

    if (idxFolio === -1) {
      console.error('❌ No se encontró la columna FOLIO');
      return null;
    }

    var folioBusqueda = folio.toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var folioActual = (datos[i][idxFolio] || '').toString().toUpperCase().trim();
      if (folioActual === folioBusqueda) {
        var registro = {};
        for (var j = 0; j < headers.length; j++) {
          var valor = datos[i][j];
          // 💡 SOLUCIÓN: Formato con soporte de tiempo completo para evitar truncar datos de auditoría
          if (valor instanceof Date) {
            valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          }
          registro[headers[j]] = valor;
        }
        return registro;
      }
    }

    return null;

  } catch (err) {
    console.error('❌ Error en _obtenerRegistroPorFolio:', err);
    return null;
  }
}

/**
 * Obtiene los archivos de imagen de la subcarpeta del folio en Drive.
 * Devuelve array de { nombre, base64, tipo }
 */
function _obtenerFotosDeCarpeta(folderUrl) {
  if (!folderUrl) return [];
  try {
    // Extraer el ID de la URL de Drive
    const match = folderUrl.match(/[-\w]{25,}/);
    if (!match) return [];
    const folder = DriveApp.getFolderById(match[0]);
    const files = folder.getFiles();
    const fotos = [];

    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      if (!mime.startsWith('image/')) continue;

      const bytes = file.getBlob().getBytes();
      const base64 = Utilities.base64Encode(bytes);
      fotos.push({
        nombre: file.getName(),
        base64: base64,
        tipo: mime,
      });

      // Máximo 12 fotos en el PDF para mantenerlo ligero
      if (fotos.length >= 12) break;
    }
    return fotos;
  } catch (err) {
    console.warn('_obtenerFotosDeCarpeta:', err.message);
    return [];
  }
}




// ════════════════════════════════════════════════════════════
// UTILIDADES
// ════════════════════════════════════════════════════════════

/**
 * Formatea una fecha para mostrar en la interfaz
 * @param {string|Date} val - Fecha a formatear
 * @returns {string} Fecha formateada (dd/MM/yyyy)
 */
function _formatearFecha(val) {
  if (!val) return 'No disponible';
  
  // Si es un string, convertirlo a Date
  var d = typeof val === 'string' ? new Date(val) : val;
  
  // Verificar si es una fecha válida
  if (!(d instanceof Date) || isNaN(d.getTime())) {
    return String(val);
  }
  
  // Formatear como dd/MM/yyyy
  var dia = String(d.getDate()).padStart(2, '0');
  var mes = String(d.getMonth() + 1).padStart(2, '0');
  var anio = d.getFullYear();
  
  return dia + '/' + mes + '/' + anio;
}
/**
 * Formatea una fecha con hora para mostrar en la interfaz
 * @param {string|Date} val - Fecha a formatear
 * @returns {string} Fecha formateada (dd/MM/yyyy HH:mm)
 */
function _formatearFechaHora(val) {
  if (!val) return 'No disponible';
  
  var d = typeof val === 'string' ? new Date(val) : val;
  if (!(d instanceof Date) || isNaN(d.getTime())) {
    return String(val);
  }
  
  var dia = String(d.getDate()).padStart(2, '0');
  var mes = String(d.getMonth() + 1).padStart(2, '0');
  var anio = d.getFullYear();
  var horas = String(d.getHours()).padStart(2, '0');
  var minutos = String(d.getMinutes()).padStart(2, '0');
  
  return dia + '/' + mes + '/' + anio + ' ' + horas + ':' + minutos;
}
// ============================================================
// 16. GENERACIÓN Y DESCARGA DE PDF
// ============================================================

/**
 * Genera un PDF para descarga directa desde la interfaz web
 * @param {string} token - Token de sesión
 * @param {string} folio - Folio del registro
 * @returns {Object} { ok, pdfBlob, fileName, error }
 */
function generarPDFParaDescarga(token, folio) {
  try {
    // Usar la función base
    const resultado = _generarPDFBase(token, folio, {
      subirADrive: false,          // No subir a Drive para descarga directa
      actualizarBitacora: false,   // No actualizar bitácora
      registrarAuditoria: true     // Registrar auditoría
    });

    if (!resultado.success) {
      return { 
        ok: false, 
        error: resultado.error 
      };
    }

    // Retornar para descarga
    return {
      ok: true,
      pdfBlob: resultado.pdfBlob,
      fileName: `Orden_Servicio_${folio}.pdf`,
      folio: folio
    };

  } catch (err) {
    console.error('❌ Error en generarPDFParaDescarga:', err);
    return { 
      ok: false, 
      error: 'Error al generar PDF: ' + err.message 
    };
  }
}


/**
 * Genera el HTML para el reporte (Diseño limpio, profesional y 100% compatible con Google PDF)
 * @param {Object} registro - Datos del registro
 * @param {Array} fotos - Array de fotos { nombre, base64, tipo }
 * @returns {string} HTML completo
 */
function _generarHTMLReporte(registro, fotos) {
  try {
    // ── VALORES POR DEFECTO TOLERANTES ──
    var folio = registro.FOLIO || 'N/A';
    var economico = registro.ECONOMICO || 'N/A';
    var placas = registro.PLACAS || 'N/A';
    var serieGPS = registro.SERIE_GPS || 'N/A';
    var tipoServicio = registro.TIPO_REVISION || registro.TIPO_SERVICIO || 'N/A';
    var plataforma = registro.PLATAFORMA || 'SAMSARA';
    var fechaServicio = registro.FECHA_SERVICIO || 'N/A';
    var tecnicoNombre = registro.NOMBRE_TECNICO || registro.TECNICO_NOMBRE || 'N/A';
    var tecnicoId = registro.TECNICO_ID || 'N/A';
    var detalle = registro.DETALLE_TRABAJO || 'Sin detalles registrados';
    var estado = registro.ESTADO || 'N/A';
    var precio = parseFloat(registro.PRECIO_UNITARIO || 0);
    if (isNaN(precio)) precio = 0;

    var fechaEmision = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    // ── CONFIGURACIÓN DE ESTADO BADGE ──
    var badgeColor = '#6c757d';
    var estadoUpper = estado.toUpperCase().trim();
    if (estadoUpper === 'BORRADOR') badgeColor = '#ffc107';
    else if (estadoUpper === 'LISTO PARA PAGO') badgeColor = '#17a2b8';
    else if (estadoUpper === 'PAGADO') badgeColor = '#28a745';
    else if (estadoUpper === 'APROBADO') badgeColor = '#1a73e8';

    // ── HOMOLOGACIÓN DEL TIPO DE SERVICIO ──
    var tipoUpper = tipoServicio.toUpperCase().trim();
    var badgeService = tipoServicio;
    if (tipoUpper.indexOf('INSTALACION') !== -1) badgeService = 'INSTALACIÓN';
    else if (tipoUpper.indexOf('DESINSTALACION') !== -1) badgeService = 'DESINSTALACIÓN';
    else if (tipoUpper.indexOf('REVISION') !== -1) badgeService = 'REVISIÓN';
    else if (tipoUpper.indexOf('REEMPLAZO') !== -1) badgeService = 'REEMPLAZO';

    // ── GENERAR GRID DE FOTOS ──
    var gridFotos = '';
    if (fotos && fotos.length > 0) {
      var fotosHtml = '';
      // Limitar a 12 fotos para no sobrecargar el PDF
      var maxFotos = Math.min(fotos.length, 12);
      for (var f = 0; f < maxFotos; f++) {
        var foto = fotos[f];
        var base64Data = foto.base64 || '';
        if (base64Data) {
          fotosHtml += `
            <div style="display:inline-block;width:30%;margin:5px;text-align:center;">
              <img src="data:${foto.tipo || 'image/jpeg'};base64,${base64Data}" 
                   style="width:100%;height:150px;object-fit:cover;border-radius:4px;border:1px solid #e2e8f0;">
              <div style="font-size:7pt;color:#94a3b8;text-align:center;margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
                ${foto.nombre || 'Evidencia'}
              </div>
            </div>
          `;
        }
      }
      gridFotos = `
        <div class="seccion">
          <div class="seccion-titulo">📷 Evidencia Fotográfica</div>
          <div style="text-align:center;">
            ${fotosHtml}
          </div>
          <div style="font-size:7pt;color:#94a3b8;text-align:center;margin-top:4px;">
            ${fotos.length > 12 ? '* Se muestran las primeras 12 fotos' : ''}
          </div>
        </div>
      `;
    } else {
      gridFotos = `
        <div class="seccion">
          <div class="seccion-titulo">📷 Evidencia Fotográfica</div>
          <div style="color:#94a3b8; font-size:9pt; padding:4px 0;">
            <i>Sin evidencia fotográfica adjunta.</i>
          </div>
        </div>
      `;
    }

    // ── CONSTRUIR HTML COMPLETO ──
    var html = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Reporte Técnico ${folio}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 10pt;
      color: #1a202c;
      padding: 20px 25px;
      line-height: 1.4;
    }
    
    .tabla-layout {
      width: 100%;
      border-collapse: collapse;
      border: none;
      margin-bottom: 15px;
    }
    .tabla-layout td {
      border: none;
      vertical-align: top;
      padding: 4px 0;
    }
    
    .header-border {
      border-bottom: 3px solid #1a56db;
      padding-bottom: 8px;
      margin-bottom: 15px;
    }
    .logo {
      font-size: 18pt;
      font-weight: 700;
      color: #1a56db;
    }
    .sub {
      font-size: 8.5pt;
      color: #64748b;
    }
    .box-folio {
      text-align: right;
      background: #0f172a;
      color: #38bdf8;
      padding: 8px 14px;
      border-radius: 6px;
      display: inline-block;
    }
    .folio-num {
      font-size: 15pt;
      font-weight: 700;
      font-family: 'Courier New', monospace;
    }
    .folio-label {
      font-size: 7.5pt;
      color: #94a3b8;
      text-transform: uppercase;
    }
    
    .badge {
      display: inline-block;
      padding: 3px 12px;
      border-radius: 12px;
      font-size: 8.5pt;
      font-weight: 600;
      text-transform: uppercase;
    }
    .badge-estado {
      background: ${badgeColor};
      color: #fff;
    }
    .badge-servicio {
      background: #dbeafe;
      color: #1e40af;
      border: 1px solid #93c5fd;
      margin-left: 5px;
    }
    
    .seccion {
      margin-bottom: 12px;
      border: 1px solid #e2e8f0;
      border-radius: 6px;
      padding: 10px 14px;
      background: #f8fafc;
    }
    .seccion-titulo {
      font-size: 8.5pt;
      font-weight: 700;
      text-transform: uppercase;
      color: #475569;
      border-bottom: 1px solid #e2e8f0;
      padding-bottom: 4px;
      margin-bottom: 6px;
    }
    
    .campo-label {
      font-size: 7.5pt;
      color: #94a3b8;
      text-transform: uppercase;
      font-weight: 600;
    }
    .campo-valor {
      font-size: 9.5pt;
      font-weight: 500;
      color: #0f172a;
    }
    .campo-valor.mono {
      font-family: 'Courier New', monospace;
      font-weight: 600;
      color: #1a56db;
    }
    
    .detalle-box {
      background: #fff;
      border-left: 3px solid #1a56db;
      padding: 8px 12px;
      font-size: 9.5pt;
      line-height: 1.5;
      color: #334155;
      white-space: pre-wrap;
    }
    
    .tabla-costos {
      width: 100%;
      border-collapse: collapse;
      font-size: 9.5pt;
      margin-top: 4px;
      background: white;
    }
    .tabla-costos th {
      background: #0f172a;
      color: #fff;
      padding: 6px 10px;
      text-align: left;
      font-size: 8pt;
      text-transform: uppercase;
    }
    .tabla-costos td {
      padding: 6px 10px;
      border-bottom: 1px solid #e2e8f0;
    }
    .tabla-costos .total td {
      font-weight: 700;
      background: #eff6ff;
      color: #1a56db;
      border-top: 2px solid #1a56db;
    }
    
    .tabla-firmas {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
    }
    .tabla-firmas td {
      width: 33.3%;
      text-align: center;
      padding: 0 10px;
    }
    .linea-firma {
      border-top: 1.5px solid #334155;
      padding-top: 4px;
      margin-top: 35px;
    }
    .firma-box .nombre {
      font-size: 8.5pt;
      font-weight: 600;
    }
    .firma-box .cargo {
      font-size: 7.5pt;
      color: #94a3b8;
    }
    
    .footer {
      margin-top: 15px;
      border-top: 1px solid #e2e8f0;
      padding-top: 8px;
      font-size: 7.5pt;
      color: #94a3b8;
    }
  </style>
</head>
<body>

  <!-- ─── ENCABEZADO ─── -->
  <table class="tabla-layout header-border">
    <tr>
      <td>
        <div class="logo">🚛 Fleet Manager</div>
        <div class="sub">Sistema de Gestión de Flotas · Samsara</div>
        <div class="sub" style="margin-top:2px; font-size:8pt;">
          Fecha de emisión: ${fechaEmision}
        </div>
      </td>
      <td style="text-align: right;">
        <div class="box-folio">
          <div class="folio-num">${folio}</div>
          <div class="folio-label">Orden de Servicio Técnico</div>
        </div>
      </td>
    </tr>
  </table>

  <!-- ─── INDICADORES BADGES ─── -->
  <div style="margin-bottom:12px;">
    <span class="badge badge-estado">${estado}</span>
    <span class="badge badge-servicio">${badgeService}</span>
  </div>

  <!-- ─── DATOS DEL VEHÍCULO ─── -->
  <div class="seccion">
    <div class="seccion-titulo">🚗 Datos del Vehículo y Dispositivo</div>
    <table class="tabla-layout" style="margin-bottom:0;">
      <tr>
        <td style="width: 50%;">
          <div class="campo-label">Económico</div>
          <div class="campo-valor">${economico}</div>
        </td>
        <td style="width: 50%;">
          <div class="campo-label">Placas</div>
          <div class="campo-valor">${placas}</div>
        </td>
      </tr>
      <tr>
        <td>
          <div class="campo-label">Serie del GPS / Gateway</div>
          <div class="campo-valor mono">${serieGPS}</div>
        </td>
        <td>
          <div class="campo-label">Plataforma</div>
          <div class="campo-valor">${plataforma}</div>
        </td>
      </tr>
      <tr>
        <td>
          <div class="campo-label">Fecha del Servicio</div>
          <div class="campo-valor">${fechaServicio}</div>
        </td>
        <td>
          <div class="campo-label">Tipo de Revisión</div>
          <div class="campo-valor">${badgeService}</div>
        </td>
      </tr>
    </table>
  </div>

  <!-- ─── TÉCNICO RESPONSABLE ─── -->
  <div class="seccion">
    <div class="seccion-titulo">👤 Técnico Responsable</div>
    <table class="tabla-layout" style="margin-bottom:0;">
      <tr>
        <td style="width: 50%;">
          <div class="campo-label">ID Técnico</div>
          <div class="campo-valor mono">${tecnicoId}</div>
        </td>
        <td style="width: 50%;">
          <div class="campo-label">Nombre del Técnico</div>
          <div class="campo-valor">${tecnicoNombre}</div>
        </td>
      </tr>
    </table>
  </div>

  <!-- ─── DETALLE DEL TRABAJO ─── -->
  <div class="seccion">
    <div class="seccion-titulo">📋 Detalle del Trabajo Realizado</div>
    <div class="detalle-box">${detalle.replace(/\n/g, '<br>')}</div>
  </div>

  <!-- ─── EVIDENCIA FOTOGRÁFICA ─── -->
  ${gridFotos}

  <!-- ─── FIRMAS ─── -->
  <div class="seccion">
    <div class="seccion-titulo">✍️ Firmas de Autorización</div>
    <table class="tabla-firmas">
      <tr>
        <td>
          <div class="linea-firma">
            <div class="nombre">${tecnicoNombre}</div>
            <div class="cargo">Técnico Responsable</div>
          </div>
        </td>
        <td>
          <div class="linea-firma">
            <div class="nombre">_______________</div>
            <div class="cargo">Revisor / Autorizó</div>
          </div>
        </td>
        <td>
          <div class="linea-firma">
            <div class="nombre">_______________</div>
            <div class="cargo">Gerencia General</div>
          </div>
        </td>
      </tr>
    </table>
  </div>

  <!-- ─── PIE ─── -->
  <div class="footer">
    <span>Folio ${folio} · Generado el ${fechaEmision}</span>
    <span>Documento de control interno · Confidencial</span>
  </div>

</body>
</html>`;

    return html;

  } catch (err) {
    console.error('❌ Error en _generarHTMLReporte:', err);
    return null;
  }
}

/**
 * Convierte un contenido de texto HTML a formato PDF binario Blob de alta velocidad.
 * Procesa la información 100% en la memoria RAM del servidor evitando generar archivos basura en Drive.
 * 
 * @param {string} htmlContent - Código HTML completo estructurado
 * @param {string} folio - Código del folio identificador
 * @returns {Blob|null} Archivo binario PDF listo para ser descargado o adjuntado
 */
function _convertirHTMLaPDF(htmlContent, folio) {
  try {
    console.log('🖨️ Convirtiendo HTML a PDF en memoria RAM de alta velocidad...');

    if (!htmlContent) {
      console.error('❌ Contenido HTML recibido vacío.');
      return null;
    }

    // 💡 OPTIMIZACIÓN CENTRAL: Creamos el binario directamente en la memoria del script.
    // Esto evita viajes lentos a la nube de Drive y protege tu cuota de almacenamiento.
    var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'temp_' + folio + '.html');

    // El motor interno de Google Workspace renderiza el PDF conservando estilos modernos como flexbox y bordes redondeados.
    var pdfBlob = htmlBlob.getAs('application/pdf');

    // Renombrar el archivo destino con una nomenclatura limpia y estandarizada de auditoría
    var nombrePDF = 'Reporte_' + folio + '_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.pdf';
    pdfBlob.setName(nombrePDF);

    console.log('✅ PDF binario generado de forma síncrona exitosamente:', nombrePDF);
    return pdfBlob;

  } catch (err) {
    console.error('❌ Error crítico en _convertirHTMLaPDF en memoria:', err.message);

    // 💡 FALLBACK INTEGRAL: Si el motor de renderizado asíncrono llega a saturarse, 
    // ejecutamos un escape de contingencia con nomenclatura plana para no congelar la pantalla del técnico.
    try {
      console.warn('⚠️ Activando pasarela de escape alternativa para el PDF...');
      var fallbackBlob = Utilities.newBlob(htmlContent || '<h1>Reporte Técnico</h1>', 'text/html', 'fallback.html');
      var pdfFallback = fallbackBlob.getAs('application/pdf');

      var nombreFallback = 'Reporte_' + (folio || 'S-F') + '_E_FALLBACK.pdf';
      pdfFallback.setName(nombreFallback);

      return pdfFallback;
    } catch (e2) {
      console.error('❌ La pasarela de escape alternativa también colapsó:', e2.message);
      return null;
    }
  }
}
/**
 * Genera un PDF y lo envía por correo a finanzas
 * @param {string} token - Token de sesión
 * @param {string} folio - Folio del registro
 * @returns {Object} { ok, pdfUrl, correoEnviado, error }
 */
function _generarPDFBase(token, folio, opciones = {}) {
  try {
    // ============================================================
    // 1. VALIDAR SESIÓN
    // ============================================================
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      return { 
        success: false, 
        error: sesionResp.error || 'Sesión inválida o expirada' 
      };
    }

    const sesion = sesionResp.sesion;

    // ============================================================
    // 2. OBTENER REGISTRO
    // ============================================================
    const registro = _obtenerRegistroPorFolio(folio);
    if (!registro) {
      return { 
        success: false, 
        error: 'No se encontró el registro con folio: ' + folio 
      };
    }

    // ============================================================
    // 3. VERIFICAR PERMISOS
    // ============================================================
    // Solo administradores, revisores o el técnico que creó el reporte
    const esAdmin = sesion.rol === 1 || sesion.rol === 'Administrador';
    const esRevisor = sesion.rol === 2 || sesion.rol === 'Revisor';
    const esTecnico = sesion.rol === 3 || sesion.rol === 'Técnico';
    const esCreador = registro.TECNICO_ID === sesion.usuarioId || 
                      registro.TECNICO_NOMBRE === sesion.nombre;

    if (!esAdmin && !esRevisor && !(esTecnico && esCreador)) {
      return { 
        success: false, 
        error: 'No tienes permisos para generar este PDF' 
      };
    }

    // ============================================================
    // 4. OBTENER FOTOS
    // ============================================================
    let fotosUrls = [];
    if (registro.FOTOS_DRIVE_URL) {
      fotosUrls = _obtenerFotosDeCarpeta(registro.FOTOS_DRIVE_URL);
    }

    // ============================================================
    // 5. GENERAR HTML
    // ============================================================
    const html = _generarHTMLReporte(registro, fotosUrls);

    // ============================================================
    // 6. CONVERTIR A PDF
    // ============================================================
    const pdfBlob = _convertirHTMLaPDF(html, folio);

    // ============================================================
    // 7. SUBIR A DRIVE (si se solicita)
    // ============================================================
    let pdfUrl = null;
    if (opciones.subirADrive !== false) {
      try {
        const folder = _obtenerOCrearCarpetaReportes();
        const file = folder.createFile(pdfBlob);
        pdfUrl = file.getUrl();
        console.log('📄 PDF subido a Drive:', pdfUrl);
      } catch (err) {
        console.warn('⚠️ No se pudo subir PDF a Drive:', err.message);
        // No fallamos, solo continuamos sin URL
      }
    }

    // ============================================================
    // 8. ACTUALIZAR URL EN BITÁCORA (si se generó URL)
    // ============================================================
    if (pdfUrl && opciones.actualizarBitacora !== false) {
      try {
        _actualizarURLPDF(folio, pdfUrl);
      } catch (err) {
        console.warn('⚠️ No se pudo actualizar URL en bitácora:', err.message);
      }
    }

    // ============================================================
    // 9. REGISTRAR EN AUDITORÍA
    // ============================================================
    if (opciones.registrarAuditoria !== false) {
      try {
        _registrarAuditoria(
          sesion.email || 'Sistema',
          sesion.nombre || 'Sistema',
          'GENERAR_PDF',
          'REPORTES',
          'PDF generado para folio: ' + folio,
          folio,
          '',
          ''
        );
      } catch (err) {
        console.warn('⚠️ Error al registrar auditoría:', err.message);
      }
    }

    // ============================================================
    // 10. RETORNAR RESULTADO
    // ============================================================
    return {
      success: true,
      registro: registro,
      fotosUrls: fotosUrls,
      html: html,
      pdfBlob: pdfBlob,
      pdfUrl: pdfUrl,
      usuario: sesion
    };

  } catch (err) {
    console.error('❌ Error en _generarPDFBase:', err);
    return {
      success: false,
      error: 'Error al generar PDF: ' + err.message
    };
  }
}
