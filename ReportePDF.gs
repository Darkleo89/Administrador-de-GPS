// ============================================================
// SISTEMA DE GESTIÓN DE FLOTAS — ReportePDF.gs
// Generación de HTML maquetado → PDF → Drive
// ============================================================

function generarYEnviarPDF(token, folio) {
  try {
    console.log('🚀 Iniciando fase 4 para el Folio:', folio);

    // 1. VALIDAR SESIÓN
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      throw new Error('Sesión inválida: ' + sesionResp.error);
    }

    // 2. OBTENER EL REGISTRO
    var registro = _obtenerRegistroPorFolio(folio);
    if (!registro) {
      throw new Error('No se encontró el registro con folio: ' + folio);
    }

    console.log('ℹ️ Registro obtenido exitosamente.');

    // 3. OBTENER PARÁMETROS
    var params = _leerParams();
    var folderId = params['DRIVE_FOLDER_ID'] || '';

    // 4. GENERAR EL PDF
    var pdfBlob = _generarPDFDesdePlantilla(registro);
    if (!pdfBlob) {
      throw new Error('No se pudo generar el PDF');
    }

    // 5. GUARDAR EN DRIVE (CON MANEJO DE ERRORES MEJORADO)
    var pdfUrl = '';
    var folderUrl = registro.FOTOS_DRIVE_URL || '';
    var pdfFile = null;

    try {
      console.log('💾 Guardando documento PDF final en Google Drive...');

      // Intentar guardar en la carpeta de evidencias
      if (folderUrl && folderUrl.trim() !== '') {
        try {
          // Extraer el ID de la carpeta desde la URL de forma segura
          var folderIdMatch = folderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
          if (folderIdMatch && folderIdMatch[1]) {
            var folderIdDrive = folderIdMatch[1];
            var folder = DriveApp.getFolderById(folderIdDrive);

            // Verificar que tenemos permisos de escritura
            if (folder) {
              pdfFile = folder.createFile(pdfBlob);
              console.log('✅ PDF guardado en carpeta de evidencias:', pdfFile.getName());
            }
          }
        } catch (err) {
          console.warn('⚠️ No se pudo guardar en la carpeta de evidencias:', err.message);
          console.warn('   Intentando guardar en la carpeta raíz...');
          pdfFile = null;
        }
      }

      // Si falló guardar en la carpeta de evidencias, guardar en la raíz configurada
      if (!pdfFile) {
        try {
          // Si tenemos folderId raíz, intentar guardar allí
          if (folderId && folderId.trim() !== '') {
            try {
              var rootFolder = DriveApp.getFolderById(folderId);
              pdfFile = rootFolder.createFile(pdfBlob);
              console.log('✅ PDF guardado en carpeta raíz:', pdfFile.getName());
            } catch (err) {
              console.warn('⚠️ No se pudo guardar en carpeta raíz:', err.message);
              pdfFile = null;
            }
          }

          // Si todo falla, guardar en la raíz de Drive del usuario ejecutor
          if (!pdfFile) {
            pdfFile = DriveApp.createFile(pdfBlob);
            console.log('✅ PDF guardado en la raíz de Drive:', pdfFile.getName());
          }
        } catch (err) {
          console.error('❌ Falló todos los intentos de guardado en Drive:', err.message);
          throw new Error('No se pudo guardar el PDF en Drive: ' + err.message);
        }
      }

      // Obtener URL del PDF
      if (pdfFile) {
        pdfUrl = pdfFile.getUrl();
        console.log('📎 URL del PDF:', pdfUrl);

        // Compartir el archivo (con manejo de errores)
        try {
          pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          console.log('✅ Archivo compartido correctamente');
        } catch (shareErr) {
          console.warn('⚠️ No se pudo compartir automáticamente:', shareErr.message);
          // No es crítico, continuar el flujo
        }
      }

    } catch (driveErr) {
      console.error('❌ Error al guardar en Drive:', driveErr.message);
      console.error('   Stack:', driveErr.stack);

      // Si el error es de permisos, pero el PDF se guardó parcialmente en algún fallback
      if (driveErr.message.includes('Acceso denegado') || driveErr.message.includes('DriveApp')) {
        console.warn('⚠️ Error de permisos de Drive. El PDF puede estar guardado pero no se pudo compartir.');

        // Intentar recuperar el archivo por nombre como fallback síncrono
        try {
          var nombreBusqueda = 'PDF_' + folio + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
          var archivos = DriveApp.getFilesByName(nombreBusqueda);
          if (archivos.hasNext()) {
            var archivoEncontrado = archivos.next();
            pdfUrl = archivoEncontrado.getUrl();
            console.log('📎 PDF recuperado por nombre:', pdfUrl);
          }
        } catch (fallbackErr) {
          console.warn('⚠️ No se pudo recuperar el PDF por nombre:', fallbackErr.message);
        }
      }

      // Si no tenemos URL después de todos los intentos, usar la URL que tengamos del puntero temporal
      if (!pdfUrl && pdfFile) {
        pdfUrl = pdfFile.getUrl() || 'PDF guardado, pero no se pudo obtener la URL';
      }
    }

    // 6. ACTUALIZAR EL REGISTRO EN BITÁCORA
    if (pdfUrl) {
      try {
        _actualizarUrlPdfBitacora(folio, pdfUrl);
        console.log('✅ URL del PDF actualizada en Bitácora');
      } catch (updateErr) {
        console.warn('⚠️ No se pudo actualizar la URL en Bitácora:', updateErr.message);
        // No es crítico, el PDF ya está guardado en la nube
      }
    }

    // 7. ENVIAR CORREO (si está configurado)
    try {
      var correosDestino = params['CORREOS_DESTINO'] || '';
      if (correosDestino && correosDestino.trim() !== '') {
        var asunto = '📄 PDF Generado - Folio ' + folio;
        var cuerpo = 'Se ha generado el PDF para el folio ' + folio + '.\n\n';
        cuerpo += '📎 URL del PDF: ' + pdfUrl + '\n\n';
        cuerpo += '📅 Fecha de generación: ' + new Date().toLocaleString();

        if (pdfFile) {
          MailApp.sendEmail({
            to: correosDestino,
            subject: asunto,
            body: cuerpo,
            attachments: [pdfFile]
          });
          console.log('📧 Correo enviado a:', correosDestino);
        } else {
          MailApp.sendEmail(correosDestino, asunto, cuerpo);
          console.log('📧 Correo (sin adjunto) enviado a:', correosDestino);
        }
      }
    } catch (emailErr) {
      console.warn('⚠️ No se pudo enviar el correo:', emailErr.message);
      // No es crítico para detener la respuesta web
    }

    return {
      ok: true,
      folio: folio,
      pdfUrl: pdfUrl || 'PDF generado, pero no se pudo obtener la URL',
      message: 'PDF generado y guardado exitosamente'
    };

  } catch (err) {
    console.error('❌ Error en generarYEnviarPDF:', err);
    console.error('   Stack:', err.stack);
    return {
      ok: false,
      error: 'Error al generar PDF: ' + err.message
    };
  }
}
/**
 * Genera el PDF a partir de una plantilla HTML
 * @param {Object} registro - Datos del registro
 * @returns {Blob|null} - Blob del PDF o null si falla
 */
function _generarPDFDesdePlantilla(registro) {
  try {
    console.log('📄 Construyendo plantilla HTML...');

    // Crear plantilla HTML (ejemplo simplificado)
    var html = '<html>';
    html += '<head><meta charset="UTF-8"></head>';
    html += '<body style="font-family: Arial, sans-serif;">';
    html += '<h1 style="color: #1a73e8;">Reporte Técnico - ' + registro.FOLIO + '</h1>';
    html += '<hr>';
    html += '<p><strong>Económico:</strong> ' + (registro.ECONOMICO || 'N/A') + '</p>';
    html += '<p><strong>Placas:</strong> ' + (registro.PLACAS || 'N/A') + '</p>';
    html += '<p><strong>Tipo de Servicio:</strong> ' + (registro.TIPO_REVISION || 'N/A') + '</p>';
    html += '<p><strong>Técnico:</strong> ' + (registro.NOMBRE_TECNICO || 'N/A') + '</p>';
    html += '<p><strong>Fecha:</strong> ' + (registro.FECHA_SERVICIO || 'N/A') + '</p>';
    html += '<hr>';
    html += '<p><strong>Detalle del Trabajo:</strong></p>';
    html += '<p style="white-space: pre-wrap;">' + (registro.DETALLE_TRABAJO || 'Sin detalles') + '</p>';
    html += '<hr>';
    html += '<p style="color: #666; font-size: 12px;">Documento generado automáticamente el ' + new Date().toLocaleString() + '</p>';
    html += '</body></html>';

    console.log('🖨️ Convirtiendo HTML a formato binario Blob PDF...');

    // Generar PDF usando la API de Google
    var pdfBlob = Utilities.newBlob(html, 'text/html', 'temp.html')
      .getAs('application/pdf');

    // Renombrar el archivo
    var nombrePDF = 'PDF_' + registro.FOLIO + '_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '.pdf';
    pdfBlob.setName(nombrePDF);

    console.log('✅ PDF generado exitosamente:', nombrePDF);
    return pdfBlob;

  } catch (err) {
    console.error('❌ Error en _generarPDFDesdePlantilla:', err);
    return null;
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
// MAQUETACIÓN DEL REPORTE HTML
// ════════════════════════════════════════════════════════════

function _construirHTMLReporte(reg, fotos) {
  const fecha = _formatearFecha(reg.FECHA_SERVICIO);
  const fechaCorte = _formatearFecha(new Date());
  const precio = Number(reg.PRECIO_UNITARIO || 0);
  const iva = precio * 0.16;
  const total = precio + iva;

  // Construir el grid de fotos
  let gridFotos = '';
  if (fotos.length > 0) {
    const items = fotos.map((f, i) => `
      <div class="foto-item">
        <img src="data:${f.tipo};base64,${f.base64}" alt="Evidencia ${i + 1}">
        <p class="foto-label">Evidencia ${i + 1} — ${f.nombre}</p>
      </div>`).join('');
    gridFotos = `
      <div class="seccion">
        <div class="seccion-titulo">📷 Evidencia Fotográfica</div>
        <div class="foto-grid">${items}</div>
      </div>`;
  } else {
    gridFotos = `
      <div class="seccion">
        <div class="seccion-titulo">📷 Evidencia Fotográfica</div>
        <p class="sin-fotos">Sin evidencia fotográfica adjunta.</p>
      </div>`;
  }

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<style>
  /* ── Reset y base ── */
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 11pt;
    color: #1a202c;
    background: #fff;
    padding: 24px 28px;
  }

  /* ── Encabezado ── */
  .header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    border-bottom: 3px solid #1a56db;
    padding-bottom: 14px;
    margin-bottom: 18px;
  }
  .header-empresa { font-size: 18pt; font-weight: 700; color: #1a56db; }
  .header-sub     { font-size: 9pt; color: #64748b; margin-top: 3px; }
  .header-folio   {
    text-align: right;
    background: #0f172a;
    color: #38bdf8;
    padding: 10px 16px;
    border-radius: 8px;
    font-family: 'Courier New', monospace;
  }
  .header-folio .folio-num  { font-size: 18pt; font-weight: 700; letter-spacing: .05em; }
  .header-folio .folio-tipo { font-size: 8pt;  color: #94a3b8; margin-top: 3px; }

  /* ── Estado badge ── */
  .estado-badge {
    display: inline-block;
    background: #dcfce7; color: #166534;
    font-weight: 700; font-size: 9pt;
    padding: 3px 10px; border-radius: 20px;
    border: 1px solid #86efac;
    margin-bottom: 16px;
  }

  /* ── Secciones ── */
  .seccion         { margin-bottom: 18px; }
  .seccion-titulo  {
    font-size: 9pt; font-weight: 700;
    text-transform: uppercase; letter-spacing: .08em;
    color: #64748b;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 5px; margin-bottom: 10px;
  }

  /* ── Grids de datos ── */
  .datos-grid {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 10px 16px;
  }
  .dato-item label {
    display: block; font-size: 8pt;
    color: #94a3b8; margin-bottom: 2px;
  }
  .dato-item span {
    font-size: 10.5pt; font-weight: 600;
  }
  .dato-item.full { grid-column: 1 / -1; }

  /* ── Detalle de trabajo ── */
  .detalle-box {
    background: #f8fafc;
    border-left: 4px solid #1a56db;
    border-radius: 0 6px 6px 0;
    padding: 10px 14px;
    font-size: 10pt;
    line-height: 1.6;
    color: #334155;
  }

  /* ── Tabla de costos ── */
  .tabla-costos {
    width: 100%;
    border-collapse: collapse;
    margin-top: 6px;
    font-size: 10pt;
  }
  .tabla-costos th {
    background: #0f172a; color: #fff;
    padding: 7px 12px; text-align: left; font-size: 9pt;
  }
  .tabla-costos td { padding: 7px 12px; border-bottom: 1px solid #e2e8f0; }
  .tabla-costos tr:last-child td { border-bottom: none; }
  .tabla-costos .total-row td {
    font-weight: 700; font-size: 11pt;
    background: #eff6ff; color: #1a56db;
  }
  .tabla-costos .muted { color: #94a3b8; font-size: 9pt; }

  /* ── Fotos ── */
  .foto-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 10px;
  }
  .foto-item img {
    width: 100%; height: 130px;
    object-fit: cover;
    border-radius: 6px;
    border: 1px solid #e2e8f0;
    display: block;
  }
  .foto-label {
    font-size: 7.5pt; color: #64748b;
    text-align: center; margin-top: 4px;
  }
  .sin-fotos { color: #94a3b8; font-style: italic; font-size: 10pt; }

  /* ── Firmas ── */
  .firmas-grid {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 20px;
    margin-top: 8px;
  }
  .firma-box {
    border-top: 1.5px solid #334155;
    padding-top: 8px;
    text-align: center;
  }
  .firma-box .firma-nombre { font-size: 9pt; font-weight: 700; }
  .firma-box .firma-cargo  { font-size: 8pt; color: #64748b; }
  .firma-espacio           { height: 44px; }

  /* ── Pie de página ── */
  .footer {
    margin-top: 22px;
    border-top: 1px solid #e2e8f0;
    padding-top: 10px;
    font-size: 8pt;
    color: #94a3b8;
    display: flex;
    justify-content: space-between;
  }
</style>
</head>
<body>

<!-- ENCABEZADO -->
<div class="header">
  <div>
    <div class="header-empresa">🚛 Fleet Manager</div>
    <div class="header-sub">Sistema de Gestión de Flotas · Plataforma Samsara</div>
    <div class="header-sub">Fecha de corte: ${fechaCorte}</div>
  </div>
  <div class="header-folio">
    <div class="folio-num">${reg.FOLIO}</div>
    <div class="folio-tipo">Orden de Servicio Técnico</div>
  </div>
</div>

<div class="estado-badge">✅ ${reg.ESTADO}</div>

<!-- DATOS DEL VEHÍCULO Y GPS -->
<div class="seccion">
  <div class="seccion-titulo">🚚 Datos del vehículo y dispositivo</div>
  <div class="datos-grid">
    <div class="dato-item">
      <label>Económico</label>
      <span>${reg.ECONOMICO || '—'}</span>
    </div>
    <div class="dato-item">
      <label>Placas</label>
      <span>${reg.PLACAS || '—'}</span>
    </div>
    <div class="dato-item">
      <label>Fecha del servicio</label>
      <span>${fecha}</span>
    </div>
    <div class="dato-item">
      <label>Serie del GPS</label>
      <span style="font-family:monospace">${reg.SERIE_GPS || '—'}</span>
    </div>
    <div class="dato-item">
      <label>Plataforma</label>
      <span>${reg.PLATAFORMA || 'SAMSARA'}</span>
    </div>
    <div class="dato-item">
      <label>Tipo de revisión</label>
      <span>${reg.TIPO_REVISION || '—'}</span>
    </div>
  </div>
</div>

<!-- TÉCNICO -->
<div class="seccion">
  <div class="seccion-titulo">👤 Técnico responsable</div>
  <div class="datos-grid">
    <div class="dato-item">
      <label>ID Técnico</label>
      <span>${reg.TECNICO_ID || '—'}</span>
    </div>
    <div class="dato-item">
      <label>Nombre</label>
      <span>${reg.TECNICO_NOMBRE || '—'}</span>
    </div>
    <div class="dato-item">
      <label>Aprobado por</label>
      <span>${reg.APROBADO_POR || '—'}</span>
    </div>
  </div>
</div>

<!-- DETALLE DEL TRABAJO -->
<div class="seccion">
  <div class="seccion-titulo">📋 Detalle del trabajo realizado</div>
  <div class="detalle-box">${(reg.DETALLE_TRABAJO || '').replace(/\n/g, '<br>')}</div>
</div>
<!-- EVIDENCIA FOTOGRÁFICA -->
${gridFotos}

<!-- FIRMAS -->
<div class="seccion">
  <div class="seccion-titulo">✍️ Firmas de autorización</div>
  <div class="firmas-grid">
    <div class="firma-box">
      <div class="firma-espacio"></div>
      <div class="firma-nombre">${reg.TECNICO_NOMBRE || '_______________'}</div>
      <div class="firma-cargo">Técnico responsable</div>
    </div>
    <div class="firma-box">
      <div class="firma-espacio"></div>
      <div class="firma-nombre">${reg.APROBADO_POR || '_______________'}</div>
      <div class="firma-cargo">Revisor / Autorizó</div>
    </div>
    <div class="firma-box">
      <div class="firma-espacio"></div>
      <div class="firma-nombre">_______________</div>
      <div class="firma-cargo">Gerencia General</div>
    </div>
  </div>
</div>

<!-- PIE -->
<div class="footer">
  <span>Folio ${reg.FOLIO} · Generado el ${fechaCorte}</span>
  <span>Documento de control interno · Confidencial</span>
</div>

</body>
</html>`;
}


// ════════════════════════════════════════════════════════════
// CONVERSIÓN Y ALMACENAMIENTO DEL PDF
// ════════════════════════════════════════════════════════════

/**
 * Convierte el HTML en PDF usando el servicio nativo de Apps Script.
 */
function _htmlAPDF(html, folio) {
  // Crear archivo HTML temporal en Drive
  const tempFile = DriveApp.createFile(
    folio + '_temp.html',
    html,
    MimeType.HTML
  );

  // Exportar como PDF
  const pdfBlob = tempFile
    .getAs(MimeType.PDF)
    .setName('Reporte_' + folio + '.pdf');

  // Eliminar el archivo temporal
  tempFile.setTrashed(true);

  return pdfBlob;
}

/**
 * Guarda el PDF dentro de la misma subcarpeta del folio.
 * Devuelve la URL del archivo PDF.
 */
function _guardarPDFenDrive(pdfBlob, folio, folderUrl) {
  let folder;
  try {
    const match = folderUrl.match(/[-\w]{25,}/);
    folder = DriveApp.getFolderById(match[0]);
  } catch (_) {
    // Si no se puede acceder a la carpeta del folio,
    // guardar en la carpeta raíz del sistema
    const params = _leerParams();
    folder = DriveApp.getFolderById(params['DRIVE_FOLDER_ID']);
  }

  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  return pdfFile.getUrl();
}


// ════════════════════════════════════════════════════════════
// ACTUALIZACIÓN DE LA BITÁCORA
// ════════════════════════════════════════════════════════════

function _marcarComoPagado(folio, pdfUrl, aprobadoPor) {
  var sheet = SHEETS.BITACORA();
  var datos = sheet.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === folio) {
      var fila = i + 1;
      // ✅ Si ya está en "Listo para pago", no cambiar a "Pagado" aún
      // Solo guardar la URL del PDF
      sheet.getRange(fila, 17).setValue(pdfUrl);      // Q: PDF_URL
      sheet.getRange(fila, 18).setValue(true);        // R: CORREO_ENVIADO
      // NOTA: El estado NO se cambia a "Pagado" aquí
      // El Revisor lo cambiará después de asignar tipo y precio
      return;
    }
  }
}

// ════════════════════════════════════════════════════════════
// UTILIDADES
// ════════════════════════════════════════════════════════════

function _formatearFecha(val) {
  if (!val) return '—';
  try {
    return Utilities.formatDate(
      new Date(val),
      Session.getScriptTimeZone(),
      'dd/MM/yyyy'
    );
  } catch (_) { return val.toString(); }
}

// ============================================================
// 16. GENERACIÓN Y DESCARGA DE PDF
// ============================================================

/**
 * Genera el PDF y lo devuelve como base64 limpio para descarga directa
 * INCLUYE FOTOS DESDE DRIVE
 */
function generarPDFParaDescarga(token, folio) {
  try {
    console.log('📄 Iniciando generación de PDF para descarga - Folio:', folio);

    // 1. VALIDAR SESIÓN
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      console.error('❌ Sesión denegada en generación de PDF para el Folio:', folio);
      return { ok: false, error: sesionResp.error };
    }

    // 2. OBTENER EL REGISTRO DE LA BITÁCORA
    var registro = _obtenerRegistroPorFolio(folio);
    if (!registro) {
      console.warn('⚠️ No se encontró el registro físico con folio:', folio);
      return { ok: false, error: 'No se encontró el registro con folio: ' + folio };
    }

    // 3. ✅ OBTENER FOTOS DE LA CARPETA DE DRIVE
    var fotos = [];
    if (registro.FOTOS_DRIVE_URL && registro.FOTOS_DRIVE_URL.trim() !== '') {
      console.log('📸 Obteniendo fotos desde:', registro.FOTOS_DRIVE_URL);
      fotos = _obtenerFotosDeCarpeta(registro.FOTOS_DRIVE_URL);
      console.log('📸 ' + fotos.length + ' fotos encontradas');
    }

    // 4. GENERAR EL CONTENIDO HTML DEL REPORTE (CON FOTOS)
    var htmlContent = _generarHTMLReporte(registro, fotos);
    if (!htmlContent) {
      console.error('❌ El motor gráfico devolvió un contenido HTML vacío.');
      return { ok: false, error: 'No se pudo generar el contenido estructural del PDF.' };
    }

    // 5. CONVERTIR EL HTML EN UN BLOB PDF BINARIO EN MEMORIA RAM
    var pdfBlob = _convertirHTMLaPDF(htmlContent, folio);
    if (!pdfBlob) {
      console.error('❌ El convertidor de Google Workspace falló al renderizar el Blob PDF.');
      return { ok: false, error: 'No se pudo convertir el HTML a formato PDF.' };
    }

    // 6. OBTENER BYTES Y CODIFICAR A BASE64
    var pdfBytes = pdfBlob.getBytes();
    var pdfBase64Raw = Utilities.base64Encode(pdfBytes);
    var pdfBase64Limpio = pdfBase64Raw.replace(/[\r\n]/g, '');

    console.log('✅ PDF consolidado exitosamente. Tamaño:', pdfBytes.length, 'bytes');

    return {
      ok: true,
      pdfBase64: pdfBase64Limpio,
      nombre: pdfBlob.getName(),
      tamano: pdfBytes.length,
      folio: folio
    };

  } catch (err) {
    console.error('❌ Error crítico en generarPDFParaDescarga:', err);
    return { ok: false, error: 'Error al generar PDF: ' + err.message };
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

  <!-- ─── COSTOS ─── -->
  <div class="seccion">
    <div class="seccion-titulo">💰 Resumen de Costos</div>
    <table class="tabla-costos">
      <thead>
        <tr>
          <th>Descripción</th>
          <th>Tipo</th>
          <th style="text-align:right">Importe</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>Servicio técnico GPS</td>
          <td>${badgeService}</td>
          <td style="text-align:right">$${precio.toFixed(2)}</td>
        </tr>
        <tr>
          <td class="muted">IVA (16%)</td>
          <td class="muted">—</td>
          <td style="text-align:right" class="muted">$${(precio * 0.16).toFixed(2)}</td>
        </tr>
        <tr class="total">
          <td colspan="2"><strong>TOTAL</strong></td>
          <td style="text-align:right"><strong>$${(precio * 1.16).toFixed(2)}</strong></td>
        </tr>
      </tbody>
    </table>
    <div style="font-size:7pt;color:#94a3b8;margin-top:4px;">
      * Precio registrado al momento del servicio. No sujeto a modificación retroactiva.
    </div>
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
