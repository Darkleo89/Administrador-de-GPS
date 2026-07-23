// ============================================================
// SISTEMA DE GESTIÓN DE FLOTAS — CorreoFinanzas.gs
// Envío automático del PDF a los destinatarios configurados
// ============================================================

/**
 * Envía el PDF del reporte por correo a todos los destinatarios
 * definidos en el parámetro CORREOS_DESTINO de la Sección C.
 * @param {Object} registro - Datos del registro
 * @param {Blob} pdfBlob - Blob del PDF
 * @param {string} pdfUrl - URL del PDF en Drive
 */
function _enviarCorreoFinanzas(registro, pdfBlob, pdfUrl) {
  try {
    const params = _leerParams();
    const destinatarios = _parsearCorreos(params['CORREOS_DESTINO'] || '');

    if (destinatarios.length === 0) {
      console.warn('⚠️ No hay destinatarios configurados en CORREOS_DESTINO.');
      return;
    }

    // ✅ Usar _formatearFecha desde code.gs (ya unificada)
    const fecha = _formatearFecha(registro.FECHA_SERVICIO);
    const precio = Number(registro.PRECIO_UNITARIO || 0);
    const total = precio * 1.16;
    const asunto = `[Fleet Manager] Orden de Servicio ${registro.FOLIO} — ${registro.TECNICO_NOMBRE}`;

    const cuerpoHtml = `
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:11pt;color:#1a202c;max-width:600px;margin:0 auto;padding:20px">

  <!-- Cabecera -->
  <div style="background:#0f172a;border-radius:10px 10px 0 0;padding:20px 24px;">
    <span style="color:#38bdf8;font-size:16pt;font-weight:700;font-family:monospace">
      🚛 Fleet Manager
    </span>
    <p style="color:#94a3b8;margin:4px 0 0;font-size:9pt">
      Sistema de Gestión de Flotas · Plataforma Samsara
    </p>
  </div>

  <!-- Cuerpo -->
  <div style="border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px;padding:24px">

    <p style="margin:0 0 16px">
      Se adjunta la <strong>Orden de Servicio Técnico</strong> lista para proceso de pago.
    </p>

    <!-- Tarjeta del folio -->
    <div style="background:#eff6ff;border-left:4px solid #1a56db;border-radius:0 8px 8px 0;
                padding:14px 18px;margin-bottom:20px">
      <div style="font-family:monospace;font-size:16pt;font-weight:700;color:#1a56db">
        ${registro.FOLIO}
      </div>
      <div style="font-size:9pt;color:#64748b;margin-top:3px">Orden de servicio aprobada</div>
    </div>

    <!-- Tabla de datos -->
    <table style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:20px">
      <tr style="background:#f8fafc">
        <td style="padding:8px 12px;color:#64748b;width:40%">Técnico</td>
        <td style="padding:8px 12px;font-weight:600">${registro.TECNICO_NOMBRE}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Vehículo (Económico)</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0;font-weight:600">${registro.ECONOMICO}</td>
      </tr>
      <tr style="background:#f8fafc">
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Placas</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0">${registro.PLACAS}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Serie GPS</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0;font-family:monospace">${registro.SERIE_GPS}</td>
      </tr>
      <tr style="background:#f8fafc">
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Tipo de revisión</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0">${registro.TIPO_REVISION}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Fecha del servicio</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0">${fecha}</td>
      </tr>
      <tr style="background:#f8fafc">
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Aprobado por</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0">${registro.APROBADO_POR || '—'}</td>
      </tr>
      ${registro.FECHA_POSIBLE_PAGO ? `
      <tr>
        <td style="padding:8px 12px;color:#64748b;border-top:1px solid #e2e8f0">Posible fecha de pago</td>
        <td style="padding:8px 12px;border-top:1px solid #e2e8f0;font-weight:600;color:#166534">
          ${_formatearFecha(registro.FECHA_POSIBLE_PAGO)}
        </td>
      </tr>` : ''}
    </table>

    <!-- Total destacado -->
    <div style="background:#0f172a;border-radius:8px;padding:14px 18px;
                display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
      <span style="color:#94a3b8;font-size:9pt">TOTAL CON IVA (16%)</span>
      <span style="color:#38bdf8;font-size:16pt;font-weight:700;font-family:monospace">
        $${total.toFixed(2)}
      </span>
    </div>

    <!-- Botón Drive -->
    <div style="text-align:center;margin-bottom:20px">
      <a href="${pdfUrl}"
         style="background:#1a56db;color:#fff;text-decoration:none;
                padding:10px 24px;border-radius:8px;font-weight:600;font-size:10pt;
                display:inline-block">
        📄 Ver PDF en Google Drive
      </a>
    </div>

    <p style="font-size:8.5pt;color:#94a3b8;border-top:1px solid #e2e8f0;padding-top:12px;margin:0">
      Este correo fue generado automáticamente por Fleet Manager.<br>
      El precio registrado ($${precio.toFixed(2)}) es un valor histórico fijo y no
      está sujeto a modificación retroactiva.
    </p>
  </div>

</body>
</html>`;

    // ✅ ENVIAR A CADA DESTINATARIO CON MANEJO DE ERRORES
    var errores = [];
    var enviados = 0;
    
    destinatarios.forEach(function(correo) {
      try {
        GmailApp.sendEmail(correo, asunto, '', {
          htmlBody    : cuerpoHtml,
          attachments : [pdfBlob],
          name        : 'Fleet Manager · Notificaciones',
        });
        enviados++;
        console.log('📧 Correo enviado a:', correo);
      } catch (err) {
        console.error('❌ Error enviando a ' + correo + ':', err.message);
        errores.push(correo + ': ' + err.message);
      }
    });

    // ✅ REGISTRAR EN AUDITORÍA
    if (enviados > 0 && typeof _registrarAuditoria === 'function') {
      try {
        _registrarAuditoria(
          'SISTEMA',
          'Sistema',
          'ENVIO_CORREO',
          'CORREOS',
          'Correo enviado a ' + enviados + ' destinatarios para folio ' + registro.FOLIO,
          registro.FOLIO,
          '',
          ''
        );
      } catch (auditErr) {
        console.warn('⚠️ Error al registrar auditoría:', auditErr.message);
      }
    }

    if (errores.length > 0) {
      console.warn('⚠️ Errores en envío de correos:', errores.join(', '));
    }

    console.log('📧 Resumen: ' + enviados + ' enviados, ' + errores.length + ' errores');
    
    return {
      success: true,
      enviados: enviados,
      errores: errores,
      total: destinatarios.length
    };

  } catch (err) {
    console.error('❌ Error en _enviarCorreoFinanzas:', err);
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Parsea el campo CORREOS_DESTINO separado por comas.
 * Limpia espacios y filtra entradas vacías o inválidas.
 * ✅ Validación mejorada
 * @param {string} cadena - Cadena de correos separados por comas
 * @returns {Array} - Array de correos válidos
 */
function _parsearCorreos(cadena) {
  if (!cadena || cadena.toString().trim() === '') {
    return [];
  }
  
  return cadena
    .toString()
    .split(',')
    .map(function(c) { return c.trim(); })
    .filter(function(c) { 
      // ✅ Validación más robusta
      return c.length > 5 && c.includes('@') && c.includes('.');
    });
}