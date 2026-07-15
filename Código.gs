// ============================================================
// SISTEMA DE GESTIÓN DE FLOTAS — Code.gs
// Google Apps Script · Backend Principal
// ============================================================

// ── CONSTANTES GLOBALES ──────────────────────────────────────
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  PARAMETROS: () => SS.getSheetByName('⚙️_Parametros'),
  TARIFAS: () => SS.getSheetByName('💰_Tarifas'),
  USUARIOS: () => SS.getSheetByName('👤_Usuarios'),
  INVENTARIO: () => SS.getSheetByName('📦_Inventario_GPS'),
  FLOTILLA: () => SS.getSheetByName('🚚_Flotilla_Fallas'),
  BITACORA: () => SS.getSheetByName('📝_Bitacora_Revisiones'),
  CONSULTA: () => SS.getSheetByName('📊_Consulta_Tecnicos'),
  TIPOS_EQUIPO: () => SS.getSheetByName('📋_Tipos_Equipo'),
  ACCESORIOS: () => SS.getSheetByName('🔧_Accesorios_Stock'),
};
function _leerParams() {
  const sheet = SHEETS.PARAMETROS();
  if (!sheet) {
    console.warn('No se encontró la hoja de parámetros, usando valores por defecto');
    return {
      'FOLIO_PREFIJO': 'FS',
      'FOLIO_ULTIMO': '0',
      'DRIVE_FOLDER_ID': '',
      'CORREOS_DESTINO': '',
      'DIAS_LIMITE_DRIVE': '60',
      'COMPRESION_IMAGENES': 'true',
      'CALIDAD_IMAGEN': '80',
      'MAX_IMAGENES_PDF': '12'
    };
  }

  const datos = sheet.getDataRange().getValues();
  const mapa = {};
  // Saltar encabezado (fila 1)
  for (var i = 1; i < datos.length; i++) {
    var clave = datos[i][0];
    var valor = datos[i][1];
    if (clave) mapa[clave.toString()] = valor ? valor.toString() : '';
  }
  return mapa;
}

// Fila de inicio de datos en cada sección de Configuración
const CFG_ROWS = {
  TARIFAS: 2,   // Sección A empieza en fila 2
  USUARIOS: 20,  // Sección B empieza en fila 20
  PARAMS: 40,  // Sección C empieza en fila 40
};

const ESTADO = {
  BORRADOR: 'Borrador',
  LISTO_PAGO: 'Listo para pago',
  PAGADO: 'Pagado',
};


// ════════════════════════════════════════════════════════════
// 1. PUNTO DE ENTRADA DE LA WEB APP
// ════════════════════════════════════════════════════════════

function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Fleet Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Incluir archivos parciales (CSS, JS) desde el editor de Apps Script
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ════════════════════════════════════════════════════════════
// 2. AUTENTICACIÓN Y SESIÓN
// ════════════════════════════════════════════════════════════

/**
 * Verifica credenciales contra la hoja 👤_Usuarios.
 * Devuelve objeto de sesión o error.
 */
function loginUsuario(username, password) {
  try {
    console.log('=== INICIO loginUsuario ===');
    console.log('Usuario intentando:', username);

    // Usar la hoja USUARIOS en lugar de CONFIG
    const sheet = SHEETS.USUARIOS();
    if (!sheet) {
      console.error('No se encontró la hoja 👤_Usuarios');
      return { ok: false, error: 'No se encontró la hoja de usuarios. ¿La base de datos está instalada?' };
    }

    // Leer todos los datos de la hoja Usuarios
    const datos = sheet.getDataRange().getValues();
    console.log('Datos de usuarios leídos:', datos.length, 'filas');

    // Saltar encabezado (fila 1)
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var usuarioId = fila[0];
      var nombre = fila[1];
      var user = fila[2];
      var passHash = fila[3];
      var rol = fila[4];
      var activo = fila[5];

      // Saltar filas vacías
      if (!user) continue;

      console.log('Validando usuario:', user, 'activo:', activo);

      // Comparar credenciales
      var userMatch = user.toString().toLowerCase() === username.toLowerCase();
      var passMatch = passHash.toString() === hashSimple(password);

      if (userMatch && passMatch && activo === true) {
        console.log('✅ Usuario autenticado:', user);

        // Registrar último acceso
        _actualizarUltimoAcceso(user, sheet, datos, i);

        // Crear token de sesión firmado
        var sesion = {
          usuarioId: usuarioId.toString(),
          nombre: nombre.toString(),
          username: user.toString(),
          rol: Number(rol),
          token: _generarToken(user.toString(), Number(rol)),
          timestamp: new Date().toISOString(),
        };

        // Guardar en PropertiesService (por sesión de script)
        PropertiesService.getScriptProperties()
          .setProperty('SESION_' + sesion.token, JSON.stringify(sesion));

        console.log('✅ Sesión creada para:', sesion.nombre);
        return { ok: true, sesion: sesion };
      }
    }

    console.warn('❌ Credenciales incorrectas para usuario:', username);
    return { ok: false, error: 'Credenciales incorrectas o usuario inactivo.' };

  } catch (err) {
    console.error('❌ Error en loginUsuario:', err);
    console.error('Stack:', err.stack);
    return { ok: false, error: 'Error interno: ' + err.message };
  }
}

/**
 * Hash simple reproducible (SHA-256 vía Utilities).
 * Para producción robusta, considera un salt fijo por usuario.
 */
function hashSimple(texto) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    texto,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function _generarToken(username, rol) {
  const raw = username + rol + new Date().getTime() + Math.random();
  return Utilities.base64Encode(raw).replace(/[^a-zA-Z0-9]/g, '').slice(0, 32);
}

function _actualizarUltimoAcceso(username, sheet, datos, rowIndex) {
  try {
    // rowIndex es la posición en el array (0-based)
    // La fila real en la hoja es rowIndex + 1
    sheet.getRange(rowIndex + 1, 7).setValue(new Date()); // Columna G = ULTIMO_ACCESO
  } catch (err) {
    console.warn('No se pudo actualizar último acceso:', err.message);
  }
}

/**
 * Valida que un token de sesión siga vigente.
 */
function validarSesion(token) {
  const raw = PropertiesService.getScriptProperties()
    .getProperty('SESION_' + token);
  if (!raw) return { ok: false, error: 'Sesión expirada. Inicia sesión de nuevo.' };
  return { ok: true, sesion: JSON.parse(raw) };
}

function cerrarSesion(token) {
  PropertiesService.getScriptProperties().deleteProperty('SESION_' + token);
  return { ok: true };
}


// ════════════════════════════════════════════════════════════
// 3. GENERADOR DE FOLIOS
// ════════════════════════════════════════════════════════════

/**
 * Genera y reserva el siguiente folio consecutivo.
 * Lee FOLIO_ULTIMO de Params, incrementa y guarda.
 * Ejemplo: FS-0088
 */
function generarFolio() {
  const params = _leerParams();

  const prefijo = params['FOLIO_PREFIJO'] || 'FS';
  const ultimo = parseInt(params['FOLIO_ULTIMO'] || '0', 10);
  const nuevo = ultimo + 1;
  const folio = prefijo + '-' + String(nuevo).padStart(4, '0');

  _escribirParam('FOLIO_ULTIMO', String(nuevo));

  return folio;
}

function _escribirParam(clave, valor) {
  const sheet = SHEETS.PARAMETROS();
  if (!sheet) {
    console.warn('No se encontró la hoja de parámetros');
    return;
  }

  const datos = sheet.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === clave) {
      sheet.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
}


// ════════════════════════════════════════════════════════════
// 4. RECEPCIÓN DE REPORTE (Orden de Servicio)
// ════════════════════════════════════════════════════════════

/**
 * Recibe los datos del formulario de la Web App.
 * Crea la carpeta de evidencias en Drive y escribe la fila en Bitácora.
 *
 * @param {string} token       - Token de sesión del técnico
 * @param {Object} datos       - Campos del formulario
 * @param {Array}  archivos    - [{nombre, tipo, base64}, ...]
 */
/**
 * Punto de entrada principal mejorado para recibir, procesar y guardar reportes técnicos.
 * Soporta la creación de registros nuevos y la actualización de borradores existentes.
 * 
 * @param {string} token - Token de sesión del técnico
 * @param {Object} datos - Objeto con los campos del formulario de instalación
 * @param {Array} archivos - Arreglo de fotos evidencias en Base64
 * @returns {Object} { ok: boolean, folio: string, folderUrl: string }
 */
function recibirReporte(token, datos, archivos) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  const sesion = sesionResp.sesion;

  // Validación de permisos según jerarquía de roles
  if (sesion.rol > 3) return { ok: false, error: 'Sin permisos para reportar.' };

  try {
    const sheet = SHEETS.BITACORA();
    if (!sheet) return { ok: false, error: 'No se encontró la hoja de Bitácora.' };

    const ahora = new Date();
    const datosBitacora = sheet.getDataRange().getValues();
    const headers = datosBitacora[0];

    // Localizamos los índices de las columnas clave para evitar depender de posiciones rígidas
    const idxFolio = headers.indexOf('FOLIO');
    const idxEstado = headers.indexOf('ESTADO');

    // 💡 SOLUCIÓN 3: DETECTAR SI ES EDICIÓN DE UN BORRADOR EXISTENTE
    // Si el objeto 'datos' incluye un folio de edición válido, buscaremos su fila para sobreescribirla
    let folio = datos.folioEdicion || null;
    let filaDestino = -1;
    let folderUrl = '';

    if (folio) {
      const folioBusqueda = folio.toString().toUpperCase().trim();
      for (let i = 1; i < datosBitacora.length; i++) {
        if ((datosBitacora[i][idxFolio] || '').toString().toUpperCase().trim() === folioBusqueda) {
          filaDestino = i + 1; // Fila real en Google Sheets
          folderUrl = datosBitacora[i][headers.indexOf('FOTOS_DRIVE_URL')] || ''; // Reutilizamos la carpeta original
          break;
        }
      }
    }

    // Si no se especificó un folio de edición o no se encontró en la hoja, asumimos que es un REPORTE NUEVO
    if (filaDestino === -1) {
      folio = generarFolio(); // Generamos un folio secuencial nuevo de paquete (Ej: FS-0042)
      // Disparamos la creación de la carpeta de Drive solo para reportes nuevos
      folderUrl = _subirFotosDrive(folio, archivos);
    } else {
      // Si es una edición y el técnico subió fotos nuevas, las anexamos a la carpeta existente
      if (archivos && archivos.length > 0 && folderUrl !== '') {
        _anexarFotosACarpetaExistente(folderUrl, archivos);
      }
    }

    // 💡 SOLUCIÓN 1: ASIGNACIÓN INTELIGENTE DE ESTATUS SEGÚN CHECKBOX DEL FRONTEND
    // Si el técnico marcó "Guardar como borrador", asignamos BORRADOR. 
    // Si presionó "Enviar", el registro avanza formalmente a la cola de revisión de oficina.
    const estatusFinal = datos.esBorrador ? 'Borrador' : 'Listo para pago';

    // 💡 SOLUCIÓN 2: ASIGNACIÓN COMPATIBLE DEL TIPO DE SERVICIO (Para motor PDF Fase 4)
    // Extraemos la selección real del operador (instalacion, desinstalacion, reemplazo, revision)
    const tipoServicioReal = (datos.tipoServicio || 'instalacion').toString().toUpperCase().trim();

    // Estructura completa alineada con las columnas de tu hoja de bitácora
    const nuevaFila = [
      folio,                                     // A: FOLIO
      datos.fechaServicio ? new Date(datos.fechaServicio) : ahora, // B: FECHA_SERVICIO
      (datos.economico || '').toString().trim(), // C: ECONOMICO
      (datos.placas || '').toString().trim(),    // D: PLACAS
      (datos.serieGPS || '').toString().trim(),   // E: SERIE_GPS (OCR / Manual Samsara)
      tipoServicioReal,                          // F: TIPO_REVISION / TIPO_SERVICIO (Crucial para el PDF)
      datos.plataforma || 'SAMSARA',             // G: PLATAFORMA
      sesion.usuarioId,                          // H: TECNICO_ID
      sesion.nombre,                             // I: NOMBRE_TECNICO
      datos.detalleTrabajo,                      // J: DETALLE_TRABAJO
      0,                                         // K: PRECIO_UNITARIO (Se tasa posteriormente en oficina)
      folderUrl,                                 // L: FOTOS_DRIVE_URL
      estatusFinal,                              // M: ESTADO
      '', '', '', '',                            // N, O, P, Q: Columnas complementarias vacías
      false,                                     // R: PROCESADO
      '',                                        // S: URL_PDF
      ahora                                      // T: FECHA_MODIFICACION
    ];

    // Operación de guardado en la base de datos
    if (filaDestino === -1) {
      // Inserción limpia al final si es registro nuevo
      sheet.appendRow(nuevaFila);
      console.log('✅ Nuevo reporte creado exitosamente. Folio:', folio);
    } else {
      // Sobreescritura quirúrgica del rango exacto si estamos salvando los cambios de un borrador
      sheet.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);
      console.log('✅ Borrador actualizado exitosamente sobre la fila:', filaDestino, 'Folio:', folio);
    }

    // 💡 SINCRONIZACIÓN DE HARDWARE: Si el reporte se guardó de forma definitiva, 
    // disparamos la actualización del estado del GPS en la hoja de inventario maestro.
    if (estatusFinal === 'Listo para pago' && datos.serieGPS) {
      _actualizarEstadoGPS(datos.serieGPS, datos.economico);
    }

    return { ok: true, folio: folio, folderUrl: folderUrl, tipo: datos.tipoServicio };

  } catch (err) {
    console.error('❌ Error en recibirReporte:', err);
    return { ok: false, error: 'Error al guardar el reporte en la Bitácora: ' + err.message };
  }
}


/**
 * El Revisor/Admin asigna el tipo de revisión y recalcula el precio.
 * Cambia el estado a "Listo para pago" si el tipo es válido.
 */
function asignarTipoRevision(token, folio, tipoRevision) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  if (sesionResp.sesion.rol > 2) return { ok: false, error: 'Sin permisos.' };

  if (!tipoRevision || tipoRevision === 'PENDIENTE') {
    return { ok: false, error: 'Selecciona un tipo de revisión válido.' };
  }

  try {
    const precio = _leerPrecio(tipoRevision);
    const sheet = SHEETS.BITACORA();
    const datos = sheet.getDataRange().getValues();

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString() === folio) {
        const fila = i + 1;
        sheet.getRange(fila, 6).setValue(tipoRevision);   // F: TIPO_REVISION
        sheet.getRange(fila, 11).setValue(precio);          // K: PRECIO_UNITARIO
        sheet.getRange(fila, 13).setValue(ESTADO.LISTO_PAGO); // M: ESTADO
        return { ok: true, precio };
      }
    }
    return { ok: false, error: 'Folio no encontrado.' };

  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Lee el precio vigente de la tarifa en el momento del registro.
 * Este valor se inyecta como constante en la fila — no es una fórmula.
 */
function _leerPrecio(tipoRevision) {
  const sheet = SHEETS.TARIFAS();
  if (!sheet) {
    throw new Error('No se encontró la hoja de tarifas.');
  }

  const datos = sheet.getDataRange().getValues();
  // Saltar encabezado
  for (var i = 1; i < datos.length; i++) {
    var tipo = datos[i][0];
    var precio = datos[i][2];
    var activo = datos[i][4];

    if (
      tipo.toString().toUpperCase() === tipoRevision.toString().toUpperCase() &&
      activo === true
    ) {
      return Number(precio);
    }
  }
  throw new Error('Tipo de revisión no encontrado en la matriz de tarifas: ' + tipoRevision);
}


// ════════════════════════════════════════════════════════════
// 5. GESTIÓN DE ARCHIVOS EN GOOGLE DRIVE
// ════════════════════════════════════════════════════════════

/**
 * Crea una subcarpeta por folio dentro de la carpeta principal
 * y sube cada archivo recibido en base64.
 * Devuelve la URL de la carpeta del folio.
 */
function _subirFotosDrive(folio, archivos) {
  const params = _leerParams();
  const folderId = params['DRIVE_FOLDER_ID'];
  const rootFolder = DriveApp.getFolderById(folderId);

  // Crear subcarpeta para este folio
  const subFolder = rootFolder.createFolder(folio + '_' + Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'
  ));

  // Subir cada archivo
  if (archivos && archivos.length > 0) {
    for (const archivo of archivos) {
      if (!archivo.base64 || !archivo.nombre) continue;

      const decodificado = Utilities.newBlob(
        Utilities.base64Decode(archivo.base64),
        archivo.tipo || 'image/jpeg',
        archivo.nombre
      );
      subFolder.createFile(decodificado);
    }
  }

  // Hacer la carpeta accesible a cualquier usuario del dominio (ajustar según política)
  subFolder.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return subFolder.getUrl();
}

/**
 * Actualiza el estado del GPS en Inventario al ser instalado en un económico.
 */
function _actualizarEstadoGPS(serieGPS, economico) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) {
    console.warn('No se encontró la hoja de inventario');
    return;
  }

  var datos = sheet.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === serieGPS.toString()) {
      // Columna E (índice 4): ESTADO
      sheet.getRange(i + 1, 5).setValue('Instalado');

      // Columna F (índice 5): ECONOMICO_ASIGNADO
      sheet.getRange(i + 1, 6).setValue(economico);

      // Columna G (índice 6): FECHA_INSTALACION
      sheet.getRange(i + 1, 7).setValue(new Date());

      // Columna I (índice 9): ULTIMA_ACTUALIZACION
      sheet.getRange(i + 1, 9).setValue(new Date());

      console.log('✅ Serie instalada:', serieGPS, 'en económico:', economico);
      break;
    }
  }
}


// ════════════════════════════════════════════════════════════
// 6. FUNCIONES DEL REVISOR (Nivel 2)
// ════════════════════════════════════════════════════════════

/**
 * Carga los registros filtrables para el panel del Revisor.
 * Soporta filtro por estado, técnico y rango de fechas.
 */

function obtenerRegistros(token, filtros) {
  try {
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    const sesion = sesionResp.sesion;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('📝_Bitacora_Revisiones');
    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];

    let registros = [];
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila || !fila[0]) continue;
      if (sesion.rol === 3 && fila[7]?.toString() !== sesion.usuarioId) continue;

      const reg = {};
      headers.forEach((h, idx) => {
        let valor = fila[idx];
        // FORZAR conversión segura para evitar problemas de serialización
        if (valor instanceof Date) {
          valor = valor.toISOString();
        } else if (valor === undefined || valor === null) {
          valor = '';
        } else if (typeof valor === 'boolean' || typeof valor === 'number') {
          // dejarlo tal cual, son tipos seguros
        } else {
          valor = valor.toString();
        }
        reg[h] = valor;
      });
      registros.push(reg);
    }

    return { ok: true, registros: registros, rol: sesion.rol };

  } catch (err) {
    return { ok: false, error: 'Excepcion: ' + err.message };
  }
}

/**
 * Actualiza campos editables de un registro por el Revisor.
 * Solo permite editar: DETALLE_TRABAJO, FECHA_POSIBLE_PAGO, NOTAS_REVISOR.
 */

function actualizarRegistro(token, folio, cambios) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  const { rol } = sesionResp.sesion;

  if (rol > 2) return { ok: false, error: 'Sin permisos para editar registros.' };

  const sheet = SHEETS.BITACORA();
  const datos = sheet.getDataRange().getValues();

  const COLS = {
    TIPO_REVISION: 6,  // F
    PRECIO_UNITARIO: 11, // K
    DETALLE_TRABAJO: 10, // J
    FECHA_POSIBLE_PAGO: 14, // N
    NOTAS_REVISOR: 19, // S
    ESTADO: 13, // M
  };

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === folio) {
      const fila = i + 1;

      if (cambios.detalleTrabajo !== undefined)
        sheet.getRange(fila, COLS.DETALLE_TRABAJO).setValue(cambios.detalleTrabajo);

      if (cambios.fechaPosiblePago !== undefined)
        sheet.getRange(fila, COLS.FECHA_POSIBLE_PAGO).setValue(new Date(cambios.fechaPosiblePago));

      if (cambios.notasRevisor !== undefined)
        sheet.getRange(fila, COLS.NOTAS_REVISOR).setValue(cambios.notasRevisor);

      // Si viene un tipo de revisión válido, recalcular precio y cambiar estado
      if (cambios.tipoRevision !== undefined && cambios.tipoRevision !== 'PENDIENTE') {
        try {
          const precio = _leerPrecio(cambios.tipoRevision);
          sheet.getRange(fila, COLS.TIPO_REVISION).setValue(cambios.tipoRevision);
          sheet.getRange(fila, COLS.PRECIO_UNITARIO).setValue(precio);
          sheet.getRange(fila, COLS.ESTADO).setValue(ESTADO.LISTO_PAGO);
        } catch (err) {
          return { ok: false, error: 'Tipo de revisión inválido: ' + err.message };
        }
      }

      return { ok: true };
    }
  }
  return { ok: false, error: 'Folio no encontrado: ' + folio };
}

/**
 * Aprueba un registro para firma del gerente.
 * Cambia estado a "Aprobado" e inyecta quién aprobó y cuándo.
 */
function aprobarRegistro(token, folio) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  const sesion = sesionResp.sesion;

  if (sesion.rol > 2) return { ok: false, error: 'Sin permisos para aprobar.' };

  const sheet = SHEETS.BITACORA();
  const datos = sheet.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === folio) {
      const fila = i + 1;
      sheet.getRange(fila, 13).setValue(ESTADO.APROBADO);      // M: ESTADO
      sheet.getRange(fila, 15).setValue(sesion.nombre);         // O: APROBADO_POR
      sheet.getRange(fila, 16).setValue(new Date());            // P: FECHA_APROBACION
      return { ok: true };
    }
  }
  return { ok: false, error: 'Folio no encontrado.' };
}


// ════════════════════════════════════════════════════════════
// 7. UTILIDADES PARA EL FRONTEND
// ════════════════════════════════════════════════════════════

/** Devuelve la lista de técnicos activos para el selector del Revisor. */
function obtenerTecnicos(token) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  if (sesionResp.sesion.rol > 2) return { ok: false, error: 'Sin permisos.' };

  const sheet = SHEETS.USUARIOS();
  if (!sheet) {
    return { ok: false, error: 'No se encontró la hoja de usuarios.' };
  }

  const datos = sheet.getDataRange().getValues();
  const tecnicos = [];

  // Saltar encabezado
  for (var i = 1; i < datos.length; i++) {
    var usuarioId = datos[i][0];
    var nombre = datos[i][1];
    var rol = datos[i][4];
    var activo = datos[i][5];

    if (usuarioId && activo === true && Number(rol) === 3) {
      tecnicos.push({ id: usuarioId.toString(), nombre: nombre.toString() });
    }
  }

  return { ok: true, tecnicos: tecnicos };
}

/** Devuelve catálogo de series GPS disponibles (estado = Disponible). */
function obtenerGPSDisponibles(token) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  const sheet = SHEETS.INVENTARIO();
  const datos = sheet.getDataRange().getValues();
  const disponibles = datos
    .slice(1)
    .filter(f => f[0] && f[4].toString() === 'Disponible')
    .map(f => ({ serie: f[0].toString(), modelo: f[1].toString() }));

  return { ok: true, disponibles };
}

/** Devuelve el catálogo de económicos y placas de la Flotilla activa. */
function obtenerFlotilla(token) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  const sheet = SHEETS.FLOTILLA();
  const datos = sheet.getDataRange().getValues();
  const flotilla = datos
    .slice(1)
    .filter(f => f[0] && f[8].toString() === 'Activo')
    .map(f => ({ economico: f[0].toString(), placas: f[1].toString(), tipo: f[2].toString() }));

  return { ok: true, flotilla };
}

// ════════════════════════════════════════════════════════════
// FASE 4 — Expuesta al frontend (Scripts.html)
// ════════════════════════════════════════════════════════════

/**
 * Genera PDF y envía correo para un lote de folios aprobados.
 * Permite procesar varios folios de un corte al mismo tiempo.
 *
 * @param {string}   token   - Token de sesión
 * @param {string[]} folios  - Array de folios ["FS-0041","FS-0042"]
 */
function procesarCorte(token, folios) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  if (sesionResp.sesion.rol > 2)
    return { ok: false, error: 'Sin permisos para procesar cortes.' };

  const resultados = [];
  for (const folio of folios) {
    const res = generarYEnviarPDF(token, folio);
    resultados.push({ folio, ...res });
  }

  const exitosos = resultados.filter(r => r.ok).length;
  const fallidos = resultados.filter(r => !r.ok).length;

  return {
    ok: fallidos === 0,
    resultados,
    resumen: `${exitosos} procesados correctamente, ${fallidos} con error.`,
  };
}

// Pruebas
function verHash() {
  console.log(hashSimple('laContraseñaQueQuieras'));
}

function diagnosticoCompleto() {
  try {
    const resultado = obtenerRegistros('TGVvbmVsIENhdWljaDExNzgxODk4ODYy', {});
    console.log('Resultado:', JSON.stringify(resultado));
  } catch (err) {
    console.log('ERROR CAPTURADO:', err.message);
    console.log('STACK:', err.stack);
  }
}
// ============================================================
// 10. EXPOSICIÓN DE FUNCIONES PARA LA WEB APP
// ============================================================

/**
 * Expone las funciones de verificación e instalación para el frontend.
 * Estas funciones son llamadas desde el HTML con google.script.run
 */

function obtenerEstadoBaseDatos() {
  return VerificacionBaseDatos.obtenerEstadoBaseDatos();
}

function instalarBaseDatosDesdeWeb() {
  return VerificacionBaseDatos.instalarBaseDatosDesdeWeb();
}

function verificarBaseDatos() {
  return VerificacionBaseDatos.verificarBaseDatos();
}
// ============================================================
// 11. OPERACIÓN DEL INSTALADOR - MEJORAS
// ============================================================


/**
 * Punto de entrada principal avanzado de la Bitácora (Versión Consolidada y Blindada).
 * Soporta la creación de registros nuevos, actualización quirúrgica de borradores,
 * mapeo inteligente de estatus y procesamiento síncrono de hardware Samsara.
 * 
 * @param {string} token - Token de sesión del usuario
 * @param {Object} datos - Parámetros y campos recolectados en el formulario web
 * @param {Array} archivos - Evidencias fotográficas en Base64
 * @param {Object} ubicacion - Coordenadas de geolocalización (latitud y longitud o null)
 * @returns {Object} Respuesta transaccional formal con metadatos del folio
 */
function recibirReporteMejorado(token, datos, archivos, ubicacion) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  // Filtro de control de permisos por jerarquía de roles
  if (sesion.rol > 3) return { ok: false, error: 'Sin permisos para reportar.' };

  try {
    var ahora = new Date();
    var latitud = ubicacion ? ubicacion.latitud : null;
    var longitud = ubicacion ? ubicacion.longitud : null;

    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Bitácora.' };
    }

    var datosBitacora = sheet.getDataRange().getValues();
    var headers = datosBitacora[0];

    // 💡 INDEXACIÓN ELÁSTICA: Calculamos las posiciones reales de las columnas en Sheets
    var idxFolio = headers.indexOf('FOLIO');
    var idxEstado = headers.indexOf('ESTADO');
    var idxFotosUrl = headers.indexOf('FOTOS_DRIVE_URL');

    if (idxFolio === -1 || idxEstado === -1) {
      console.error('❌ Columnas estructurales no encontradas en la Bitácora:', { idxFolio, idxEstado });
      return { ok: false, error: 'Estructura de Bitácora inválida. Contacte al administrador.' };
    }

    // 💡 DETECTOR INTELIGENTE DE SOBREESCRITURA DE BORRADORES
    var folio = datos.folioEdicion || null;
    var filaDestino = -1;
    var folderUrl = '';

    if (folio) {
      var folioBusqueda = folio.toString().toUpperCase().trim();
      for (var k = 1; k < datosBitacora.length; k++) {
        if ((datosBitacora[k][idxFolio] || '').toString().toUpperCase().trim() === folioBusqueda) {
          filaDestino = k + 1;
          folderUrl = datosBitacora[k][idxFotosUrl] || '';
          break;
        }
      }
    }

    // Si no es edición de borrador o no se halló la fila vieja, asignamos una inserción nueva limpia
    if (filaDestino === -1) {
      folio = generarFolio();
    }

    var tipoServicio = (datos.tipoServicio || 'instalacion').toString().toLowerCase().trim();

    // 💡 ESTADO ADAPTATIVO DINÁMICO
    var estado = datos.esBorrador ? 'Borrador' : 'Listo para pago';

    var gateway = datos.gateway || '';
    var camara = datos.camara || '';
    var accesorios = datos.accesorios || {};
    var detalleTrabajo = datos.detalleTrabajo || '';

    console.log('📋 Procesando Folio [' + folio + '] Tipo:', tipoServicio, 'Estado asignado:', estado);

    // ============================================================
    // CASE 1. DESINSTALACIÓN
    // ============================================================
    if (tipoServicio === 'desinstalacion') {
      console.log('🔄 Procesando DESINSTALACIÓN...');

      var serieActual = _obtenerSerieGPSporEconomico(datos.economico);
      console.log('Serie actual del vehículo:', serieActual);

      if (!serieActual) {
        return { ok: false, error: 'El vehículo ' + datos.economico + ' no tiene Gateway instalado.' };
      }

      // Si el reporte se envía de forma formal definitiva, liberamos el hardware de forma síncrona
      if (estado === 'Listo para pago') {
        _liberarSerieGPS(serieActual);
      }

      var detalleCompleto = detalleTrabajo + '\n\n📋 DESINSTALACIÓN DE DISPOSITIVOS:\n';
      detalleCompleto += '📡 Gateway VG retirado: ' + serieActual;
      if (camara) detalleCompleto += '\n📷 Cámara CM retirada: ' + camara;

      // Barrido tolerante de llaves booleanas dinámicas
      if (accesorios.ARNES || accesorios['ARNES']) detalleCompleto += '\n🔌 Arnés OBD retirado';
      if (accesorios.BOTON || accesorios['BOTON']) detalleCompleto += '\n🔴 Botón de pánico retirado';
      if (accesorios.CORTE || accesorios['CORTE']) detalleCompleto += '\n⛔ Corte de motor EI retirado';

      // Gestión controlada de evidencias en Drive
      if (archivos && archivos.length > 0) {
        if (filaDestino === -1 || folderUrl === '') {
          folderUrl = _subirFotosDrive(folio, archivos);
        } else {
          _anexarFotosACarpetaExistente(folderUrl, archivos);
        }
      }

      // Actualización de inventario físico de almacén
      if (estado === 'Listo para pago') {
        for (var tipoAccesorio in accesorios) {
          if (accesorios[tipoAccesorio] === true) {
            _actualizarStockAccesorios(tipoAccesorio, datos.economico, 'desinstalacion');
          }
        }
      }

      var nuevaFila = [
        folio, datos.fechaServicio || ahora, datos.economico, datos.placas, serieActual,
        'DESINSTALACION', datos.plataforma || 'SAMSARA', sesion.usuarioId, sesion.nombre,
        detalleCompleto, 0, folderUrl || '', estado, null, null, null, null, false, null,
        latitud, longitud, ahora
      ];

      if (filaDestino === -1) sheet.appendRow(nuevaFila);
      else sheet.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);

      return { ok: true, folio: folio, folderUrl: folderUrl || '', precio: 0, tipo: 'desinstalacion', serieRetirada: serieActual };
    }

    // ============================================================
    // CASE 2. INSTALACIÓN / REEMPLAZO (CON VALIDACIÓN PREVIA)
    // ============================================================
    if (tipoServicio === 'instalacion' || tipoServicio === 'reemplazo') {
      console.log('🔄 Procesando INSTALACIÓN / REEMPLAZO...');

      // ✅ VALIDACIÓN: Verificar que el vehículo NO tenga equipos instalados (solo para INSTALACIÓN)
      if (tipoServicio === 'instalacion') {
        var dispositivosExistentes = _obtenerTodosLosDispositivosPorEconomico(datos.economico);
        console.log('Dispositivos existentes en el vehículo:', dispositivosExistentes);

        if (dispositivosExistentes.gateway || dispositivosExistentes.camara) {
          var mensajeError = '⚠️ El vehículo ' + datos.economico + ' YA TIENE EQUIPOS INSTALADOS.\n\n';
          if (dispositivosExistentes.gateway) {
            mensajeError += '📡 Gateway: ' + dispositivosExistentes.gateway + '\n';
          }
          if (dispositivosExistentes.camara) {
            mensajeError += '📷 Cámara: ' + dispositivosExistentes.camara + '\n';
          }
          mensajeError += '\n✅ Debes seleccionar "Revisión / Diagnóstico" para reemplazar equipos.';
          console.error('❌ ' + mensajeError);
          return { ok: false, error: mensajeError };
        }
        console.log('✅ Vehículo sin equipos, instalación permitida.');
      }

      var seriesActuales = _obtenerTodasLasSeriesGPSporEconomico(datos.economico);
      console.log('Gateways actuales del vehículo:', seriesActuales);

      // ✅ LIBERAR ACCESORIOS VIEJOS ANTES DE INSTALAR NUEVOS
      if (estado === 'Listo para pago') {
        // Obtener accesorios actuales del vehículo
        var accesoriosViejos = _obtenerAccesoriosPorEconomico(datos.economico);
        console.log('Accesorios actuales del vehículo:', accesoriosViejos);

        // Liberar cada accesorio viejo
        for (var accViejo in accesoriosViejos) {
          if (accesoriosViejos[accViejo] === true) {
            console.log('🔧 Liberando accesorio viejo:', accViejo);
            _actualizarStockAccesorios(accViejo, datos.economico, 'desinstalacion');
          }
        }
      }

      // Liberar Gateway viejo si existe y es diferente
      if (gateway && seriesActuales.length > 0 && estado === 'Listo para pago') {
        for (var s = 0; s < seriesActuales.length; s++) {
          var serieVieja = seriesActuales[s];
          if (serieVieja !== gateway) {
            console.log('🔧 Liberando Gateway viejo por reemplazo:', serieVieja);
            _liberarSerieGPS(serieVieja);
          }
        }
      }

      // Compresión adaptativa de imágenes
      var archivosComprimidos = archivos;
      var params = _leerParams();
      if (params['COMPRESION_IMAGENES'] === 'true' && archivos && archivos.length > 0) {
        archivosComprimidos = _comprimirImagenes(archivos, parseInt(params['CALIDAD_IMAGEN'] || '80'));
      }

      if (archivosComprimidos && archivosComprimidos.length > 0) {
        if (filaDestino === -1 || folderUrl === '') {
          folderUrl = _subirFotosDrive(folio, archivosComprimidos);
        } else {
          _anexarFotosACarpetaExistente(folderUrl, archivosComprimidos);
        }
      }

      // Validaciones sintácticas de códigos
      if (gateway && !_validarSerieGPS(gateway)) {
        return { ok: false, error: 'Formato de Gateway inválido (XXXX-XXX-XXX).' };
      }
      if (camara && !_validarSerieGPS(camara)) {
        return { ok: false, error: 'Formato de Cámara inválido (XXXX-XXX-XXX).' };
      }

      // Verificación de disponibilidad en inventario
      if (gateway) {
        var disponibilidad = _verificarDisponibilidadSerie(gateway);
        if (!disponibilidad.disponible && disponibilidad.mensaje.indexOf(folio) === -1) {
          return { ok: false, error: disponibilidad.mensaje };
        }
      }
      if (camara) {
        var disponibilidadCamara = _verificarDisponibilidadSerie(camara);
        if (!disponibilidadCamara.disponible && disponibilidadCamara.mensaje.indexOf(folio) === -1) {
          return { ok: false, error: disponibilidadCamara.mensaje };
        }
      }

      var detalleCompleto = detalleTrabajo + '\n\n📋 REGISTRO DE DISPOSITIVOS:\n';
      detalleCompleto += '📡 Gateway VG: ' + gateway;
      if (camara) detalleCompleto += '\n📷 Cámara CM: ' + camara;
      if (accesorios.ARNES || accesorios['ARNES']) detalleCompleto += '\n🔌 Arnés OBD instalado';
      if (accesorios.BOTON || accesorios['BOTON']) detalleCompleto += '\n🔴 Botón de pánico instalado';
      if (accesorios.CORTE || accesorios['CORTE']) detalleCompleto += '\n⛔ Corte de motor EI instalado';

      // ✅ INSTALAR NUEVOS ACCESORIOS
      if (estado === 'Listo para pago') {
        for (var tipoAccesorio in accesorios) {
          if (accesorios[tipoAccesorio] === true) {
            console.log('🔧 Instalando accesorio nuevo:', tipoAccesorio);
            _actualizarStockAccesorios(tipoAccesorio, datos.economico, 'instalacion');
          }
        }
      }

      var nuevaFila = [
        folio, datos.fechaServicio || ahora, datos.economico, datos.placas, gateway,
        tipoServicio.toUpperCase(), datos.plataforma || 'SAMSARA', sesion.usuarioId, sesion.nombre,
        detalleCompleto, 0, folderUrl || '', estado, null, null, null, null, false, null,
        latitud, longitud, ahora
      ];

      if (filaDestino === -1) sheet.appendRow(nuevaFila);
      else sheet.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);

      // Amarre de estado físico en inventario
      if (estado === 'Listo para pago') {
        if (gateway) _actualizarEstadoGPS(gateway, datos.economico);
        if (camara) _actualizarEstadoGPS(camara, datos.economico);
      }

      return { ok: true, folio: folio, folderUrl: folderUrl || '', precio: 0, tipo: tipoServicio };
    }

    // ============================================================
    // CASE 3. REVISIÓN / DIAGNÓSTICO (SIN DUPLICADOS)
    // ============================================================
    if (tipoServicio === 'revision') {
      console.log('🔄 Procesando REVISIÓN (con posibilidad de reemplazo)...');

      // ✅ OBTENER TODOS LOS DISPOSITIVOS INSTALADOS EN EL VEHÍCULO
      var dispositivosActuales = _obtenerTodosLosDispositivosPorEconomico(datos.economico);
      console.log('Dispositivos actuales del vehículo:', dispositivosActuales);

      // ✅ SI HAY GATEWAY NUEVO Y ES DIFERENTE AL ACTUAL, LIBERAR EL VIEJO
      if (gateway && dispositivosActuales.gateway && dispositivosActuales.gateway !== gateway && estado === 'Listo para pago') {
        console.log('🔧 Liberando Gateway viejo por actualización:', dispositivosActuales.gateway);
        _liberarSerieGPS(dispositivosActuales.gateway);
      }

      // ✅ SI HAY CÁMARA NUEVA Y ES DIFERENTE A LA ACTUAL, LIBERAR LA VIEJA
      if (camara && dispositivosActuales.camara && dispositivosActuales.camara !== camara && estado === 'Listo para pago') {
        console.log('🔧 Liberando Cámara vieja por actualización:', dispositivosActuales.camara);
        _liberarSerieGPS(dispositivosActuales.camara);
      }

      // ✅ LIBERAR ACCESORIOS VIEJOS ANTES DE INSTALAR NUEVOS (EN REVISIÓN TAMBIÉN)
      if (estado === 'Listo para pago') {
        var accesoriosViejosRev = _obtenerAccesoriosPorEconomico(datos.economico);
        console.log('Accesorios actuales del vehículo (Revisión):', accesoriosViejosRev);

        for (var accViejoRev in accesoriosViejosRev) {
          if (accesoriosViejosRev[accViejoRev] === true) {
            console.log('🔧 Liberando accesorio viejo (Revisión):', accViejoRev);
            _actualizarStockAccesorios(accViejoRev, datos.economico, 'desinstalacion');
          }
        }
      }

      // Validar formato del Gateway (si se ingresó uno)
      if (gateway && !_validarSerieGPS(gateway)) {
        return { ok: false, error: 'Formato de Gateway inválido (XXXX-XXX-XXX).' };
      }

      // Verificar disponibilidad del Gateway (si se ingresó uno nuevo)
      if (gateway) {
        var disponibilidadNueva = _verificarDisponibilidadSerie(gateway);
        if (!disponibilidadNueva.disponible && disponibilidadNueva.mensaje.indexOf(folio) === -1) {
          return { ok: false, error: 'El Gateway no está disponible: ' + disponibilidadNueva.mensaje };
        }
      }

      // Validar formato de la Cámara (si se ingresó)
      if (camara && !_validarSerieGPS(camara)) {
        return { ok: false, error: 'Formato de Cámara inválido (XXXX-XXX-XXX).' };
      }

      // Verificar disponibilidad de la Cámara (si se ingresó una nueva)
      if (camara) {
        var disponibilidadCamaraNueva = _verificarDisponibilidadSerie(camara);
        if (!disponibilidadCamaraNueva.disponible && disponibilidadCamaraNueva.mensaje.indexOf(folio) === -1) {
          return { ok: false, error: 'La Cámara no está disponible: ' + disponibilidadCamaraNueva.mensaje };
        }
      }

      // Subir fotos a Drive
      if (archivos && archivos.length > 0) {
        if (filaDestino === -1 || folderUrl === '') {
          folderUrl = _subirFotosDrive(folio, archivos);
        } else {
          _anexarFotosACarpetaExistente(folderUrl, archivos);
        }
      }

      // ✅ CONSTRUIR DETALLE COMPLETO (UNA SOLA VEZ)
      var detalleCompleto = detalleTrabajo + '\n\n📋 REVISIÓN DE DISPOSITIVOS:\n🔍 Diagnóstico realizado.';

      // Gateway
      if (dispositivosActuales.gateway && gateway && dispositivosActuales.gateway !== gateway) {
        detalleCompleto += '\n📡 Gateway reemplazado: ' + dispositivosActuales.gateway + ' → ' + gateway;
      } else if (gateway) {
        detalleCompleto += '\n📡 Gateway verificado: ' + gateway;
      } else if (dispositivosActuales.gateway) {
        detalleCompleto += '\n📡 Gateway verificado: ' + dispositivosActuales.gateway;
      } else {
        detalleCompleto += '\n📡 Gateway: No se encontró dispositivo instalado.';
      }

      // Cámara
      if (camara && dispositivosActuales.camara && dispositivosActuales.camara !== camara) {
        detalleCompleto += '\n📷 Cámara reemplazada: ' + dispositivosActuales.camara + ' → ' + camara;
      } else if (camara) {
        detalleCompleto += '\n📷 Cámara instalada/verificada: ' + camara;
      } else if (dispositivosActuales.camara) {
        detalleCompleto += '\n📷 Cámara verificada: ' + dispositivosActuales.camara;
      } else {
        detalleCompleto += '\n📷 Cámara: No se encontró dispositivo instalado.';
      }

      // Accesorios
      if (accesorios.ARNES || accesorios['ARNES']) detalleCompleto += '\n🔌 Arnés OBD instalado';
      if (accesorios.BOTON || accesorios['BOTON']) detalleCompleto += '\n🔴 Botón de pánico instalado';
      if (accesorios.CORTE || accesorios['CORTE']) detalleCompleto += '\n⛔ Corte de motor EI instalado';

      console.log('📝 Detalle completo:', detalleCompleto);

      // ✅ INSTALAR NUEVOS ACCESORIOS (EN REVISIÓN TAMBIÉN)
      if (estado === 'Listo para pago') {
        if (gateway) {
          console.log('🔧 Instalando Gateway nuevo:', gateway);
          _actualizarEstadoGPS(gateway, datos.economico);
        }

        if (camara) {
          console.log('🔧 Instalando Cámara nueva:', camara);
          _actualizarEstadoGPS(camara, datos.economico);
        }

        for (var tipoAccesorio in accesorios) {
          if (accesorios[tipoAccesorio] === true) {
            _actualizarStockAccesorios(tipoAccesorio, datos.economico, 'instalacion');
          }
        }
      }

      // Guardar en Bitácora
      var nuevaFila = [
        folio,
        datos.fechaServicio || ahora,
        datos.economico,
        datos.placas,
        gateway || (dispositivosActuales.gateway || ''),
        'REVISION',
        datos.plataforma || 'SAMSARA',
        sesion.usuarioId,
        sesion.nombre,
        detalleCompleto,
        0,
        folderUrl || '',
        estado,
        null, null, null, null, false, null,
        latitud, longitud, ahora
      ];

      if (filaDestino === -1) {
        sheet.appendRow(nuevaFila);
      } else {
        sheet.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);
      }

      return {
        ok: true,
        folio: folio,
        folderUrl: folderUrl || '',
        precio: 0,
        tipo: 'revision',
        seriesViejas: dispositivosActuales,
        serieNueva: gateway
      };
    }

    // ============================================================
    // CASE 4. TIPO DE SERVICIO NO RECONOCIDO
    // ============================================================
    return { ok: false, error: 'Tipo de servicio no reconocido en el sistema: ' + tipoServicio };
  } catch (err) {
    console.error('❌ Error en recibirReporteMejorado:', err);
    return { ok: false, error: 'Error al procesar el reporte en la Bitácora: ' + err.message };
  }
}
/**
 * Obtiene los accesorios instalados actualmente en un vehículo
 * @param {string} economico - ID del vehículo
 * @returns {Object} - Objeto con los accesorios { tipo: true, ... }
 */
function _obtenerAccesoriosPorEconomico(economico) {
  var resultado = {};

  var sheet = SHEETS.ACCESORIOS();
  if (!sheet) return resultado;

  var datos = sheet.getDataRange().getValues();
  var economicoStr = economico.toString().toUpperCase().trim();

  for (var i = 1; i < datos.length; i++) {
    var accesorio = datos[i][0];
    if (!accesorio) continue;

    var economicoAsignado = (datos[i][5] || '').toString().toUpperCase().trim();

    if (economicoAsignado) {
      // Dividir la lista por comas y limpiar
      var listaEconomicos = economicoAsignado.split(',').map(function (e) {
        return e.trim();
      });

      // Si el vehículo está en la lista, marcar el accesorio como instalado
      if (listaEconomicos.indexOf(economicoStr) !== -1) {
        var tipo = datos[i][1] || accesorio.toString();
        resultado[tipo] = true;
      }
    }
  }

  return resultado;
}
/**
 * Obtiene el estado de una serie en el inventario (para depuración)
 * @param {string} serieGPS - Serie a verificar
 * @returns {Object} { estado: string, economico: string }
 */
function _obtenerEstadoSerie(serieGPS) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) return { estado: 'No encontrada', economico: '' };

  var datos = sheet.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === serieGPS.toString()) {
      return {
        estado: datos[i][4] || '',
        economico: datos[i][5] || ''
      };
    }
  }
  return { estado: 'No encontrada', economico: '' };
}

/**
 * Obtiene la Cámara instalada en un vehículo
 * @param {string} economico - ID del vehículo (ej. G-361)
 * @returns {string|null} - Serie de la cámara o null si no tiene
 */
function _obtenerCamaraPorEconomico(economico) {
  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      console.warn('No se encontró la hoja de inventario');
      return null;
    }

    var datos = sheet.getDataRange().getValues();
    var economicoStr = economico.toString().trim();

    for (var i = 1; i < datos.length; i++) {
      var serie = datos[i][0];
      var tipo = datos[i][1] || '';
      var estado = datos[i][4] || '';
      var economicoAsignado = datos[i][5] || '';

      if (economicoAsignado.toString().trim() === economicoStr && estado === 'Instalado') {
        var tipoStr = tipo.toString().toUpperCase();
        if (tipoStr.indexOf('CM') !== -1 || tipoStr.indexOf('CAMARA') !== -1) {
          return serie.toString();
        }
      }
    }

    return null;

  } catch (err) {
    console.error('Error en _obtenerCamaraPorEconomico:', err);
    return null;
  }
}

/**
 * Verifica la disponibilidad de una serie GPS en el inventario
 */
function _verificarDisponibilidadSerie(serieGPS) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) {
    return { disponible: false, mensaje: 'No se encontró la hoja de inventario.' };
  }

  var datos = sheet.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString().toUpperCase() === serieGPS.toUpperCase()) {
      var estado = datos[i][4] || '';

      if (estado === 'Disponible') {
        return { disponible: true, mensaje: 'Serie disponible' };
      } else if (estado === 'Instalado') {
        return { disponible: false, mensaje: 'La serie ' + serieGPS + ' ya está instalada en el económico ' + (datos[i][5] || 'N/A') };
      } else if (estado === 'Garantía') {
        return { disponible: false, mensaje: 'La serie ' + serieGPS + ' está en garantía' };
      } else if (estado === 'Baja') {
        return { disponible: false, mensaje: 'La serie ' + serieGPS + ' está dada de baja' };
      } else {
        return { disponible: false, mensaje: 'La serie ' + serieGPS + ' tiene estado: ' + estado };
      }
    }
  }

  return { disponible: false, mensaje: 'La serie ' + serieGPS + ' no existe en el inventario' };
}

/**
 * Valida el formato de la serie GPS: XXX-XXX-XXX
 */
function _validarSerieGPS(serie) {
  var regex = /^[A-Z0-9]{4}-[A-Z0-9]{3}-[A-Z0-9]{3}$/;
  return regex.test(serie);
}

/**
 * Verifica si una serie GPS está bloqueada por garantía
 */
function _estaBloqueadaPorGarantia(serieGPS) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) return false;

  var datos = sheet.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === serieGPS.toString()) {
      var ticket = datos[i][7];
      var estado = datos[i][4];
      if (ticket && ticket.toString().trim() !== '' && estado === 'Garantía') {
        return true;
      }
      return false;
    }
  }
  return false;
}

/**
 * Comprime imágenes (placeholder - la compresión real se hace en frontend)
 */
function _comprimirImagenes(archivos, calidad) {
  return archivos;
}

/**
 * Obtiene la ubicación GPS del técnico desde el frontend
 * (la captura se hace en el navegador)
 */

// ============================================================
// 12. BORRADORES SEGUROS - APARTAR EQUIPOS
// ============================================================

/**
 * Obtiene un borrador específico por folio de forma segura y dinámica
 * @param {string} token - Token de sesión del usuario
 * @param {string} folio - Código de folio a consultar
 * @returns {Object} { ok: boolean, registro: [...] }
 */
function obtenerBorrador(token, folio) {
  try {
    console.log('📋 obtenerBorrador - Inicio, folio:', folio);

    if (!folio) return { ok: false, error: 'El folio es requerido.' };

    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      console.error('❌ Sesión inválida:', sesionResp.error);
      return { ok: false, error: sesionResp.error };
    }

    var sesion = sesionResp.sesion;
    console.log('👤 Usuario:', sesion.nombre, 'ID:', sesion.usuarioId);

    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      console.error('❌ No se encontró la hoja de Bitácora');
      return { ok: false, error: 'No se encontró la hoja de Bitácora.' };
    }

    var datos = sheet.getDataRange().getValues();
    var headers = datos[0];

    // 💡 SOLUCIÓN 2: LOCALIZACIÓN DINÁMICA DE ENCABEZADOS CLAVE
    // Encontramos las posiciones reales en la hoja para tolerar inserciones o cambios de columnas en el futuro.
    var idxFolio = headers.indexOf('FOLIO');
    var idxTecnico = headers.indexOf('TECNICO_ID');

    if (idxFolio === -1 || idxTecnico === -1) {
      console.error('❌ Encabezados críticos no hallados en Bitácora:', { idxFolio, idxTecnico });
      return { ok: false, error: 'La estructura de la hoja de bitácora no es válida. Contacte soporte.' };
    }

    // Homologamos el folio buscado a texto limpio
    var folioBusquedaStr = folio.toString().toUpperCase().trim();
    var usuarioIdStr = (sesion.usuarioId || '').toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];

      if (!fila || fila.length === 0 || !fila[idxFolio]) continue;

      // 💡 SOLUCIÓN 3: COMPARACIÓN BLINDADA CONTRA CELDAS NULAS
      var folioActual = (fila[idxFolio] || '').toString().toUpperCase().trim();

      if (folioActual === folioBusquedaStr) {

        // Verificar que el borrador pertenece al técnico actual de forma tolerante
        var tecnicoId = (fila[idxTecnico] || '').toString().toUpperCase().trim();
        if (tecnicoId !== usuarioIdStr) {
          console.error('❌ Intento de acceso denegado: El borrador pertenece a otro técnico:', tecnicoId);
          return { ok: false, error: 'Acceso denegado. Este borrador pertenece a otro técnico.' };
        }

        // Construimos el objeto de respuesta dinámica mapeando todos los campos
        var reg = {};
        for (var j = 0; j < headers.length; j++) {
          var valor = fila[j];
          var nombreColumna = headers[j];

          if (!nombreColumna) continue; // Saltamos columnas vacías sin nombre

          // Convertir objetos tipo fecha de Google Sheets a cadenas de texto ISO uniformes
          if (valor instanceof Date) {
            valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          }

          // 💡 SOLUCIÓN 1: PRESERVACIÓN DE TIPOS BOOLEANOS NATIVEROS
          // Eliminamos la conversión forzada a string 'true'/'false' para que viaje limpio hacia el Front.
          if (typeof valor === 'boolean') {
            // Se mantiene como boolean nativo
          }

          // Convertir null/undefined a string vacío de manera segura
          if (valor === null || valor === undefined) {
            valor = '';
          }

          reg[nombreColumna] = valor;
        }

        console.log('✅ Borrador encontrado con éxito:', reg.FOLIO);
        console.log('  📋 Económico:', reg.ECONOMICO);

        return { ok: true, registro: reg };
      }
    }

    console.warn('⚠️ Borrador no encontrado en la base de datos:', folio);
    return { ok: false, error: 'El borrador con el folio ' + folio + ' no existe.' };

  } catch (err) {
    console.error('❌ Error en obtenerBorrador:', err);
    return { ok: false, error: err.message };
  }
}

// ============================================================
// 12. BORRADORES SEGUROS - APARTAR EQUIPOS
// ============================================================

/**
 * Lista los borradores del técnico actual de forma dinámica y segura
 * @param {string} token - Token de sesión del usuario
 * @returns {Object} { ok: boolean, borradores: [...] }
 */
function listarBorradores(token) {
  try {
    console.log('📋 listarBorradores - Inicio');

    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      console.error('❌ Sesión inválida:', sesionResp.error);
      return { ok: false, error: sesionResp.error };
    }

    var sesion = sesionResp.sesion;
    console.log('👤 Usuario:', sesion.nombre, 'ID:', sesion.usuarioId);

    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      console.error('❌ No se encontró la hoja de Bitácora');
      return { ok: false, error: 'No se encontró la hoja de Bitácora.' };
    }

    var datos = sheet.getDataRange().getValues();
    if (!datos || datos.length === 0) {
      console.warn('⚠️ No hay datos en Bitácora');
      return { ok: true, borradores: [] };
    }

    // 💡 SOLUCIÓN 2: OBTENER ÍNDICES DINÁMICOS DE COLUMNAS CLAVE
    // Buscamos en los encabezados la posición exacta de las columnas para evitar depender de índices fijos rígidos.
    var headers = datos[0];

    // Cambia los textos 'TECNICO_ID' y 'ESTADO' por el nombre exacto de tus encabezados en Sheets
    var idxTecnico = headers.indexOf('TECNICO_ID');
    var idxEstado = headers.indexOf('ESTADO');

    // Salida de seguridad en caso de que alguien haya renombrado o borrado los encabezados en la hoja
    if (idxTecnico === -1 || idxEstado === -1) {
      console.error('❌ Encabezados críticos no encontrados. Posiciones:', { idxTecnico, idxEstado });
      return { ok: false, error: 'La estructura de la bitácora ha cambiado. Contacte al administrador.' };
    }

    var borradores = [];
    var usuarioIdStr = (sesion.usuarioId || '').toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];

      // Saltamos filas vacías o nulas
      if (!fila || fila.length === 0 || !fila[0]) continue;

      // 💡 SOLUCIÓN 3: PROTECCIÓN CONTRA CRASH Y BÚSQUEDA TOLERANTE
      var tecnicoId = (fila[idxTecnico] || '').toString().toUpperCase().trim();
      var estado = (fila[idxEstado] || '').toString().toUpperCase().trim();

      // Comparamos en mayúsculas homologadas para evitar fallos por errores de captura manual
      if (tecnicoId === usuarioIdStr && estado === 'BORRADOR') {
        var reg = {};

        for (var j = 0; j < headers.length; j++) {
          var valor = fila[j];
          var nombreColumna = headers[j];

          if (!nombreColumna) continue; // Saltamos columnas sin nombre en el encabezado

          // ✅ CONVERTIR FECHAS A STRING (Excelente formato uniforme)
          if (valor instanceof Date) {
            valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          }

          // 💡 SOLUCIÓN 1: MANTENER BOOLEANOS NATIVOS
          // Quitamos la conversión a String 'true'/'false' para que el Frontend la interprete como boolean real.
          if (typeof valor === 'boolean') {
            // Se queda como boolean puro
          }

          // ✅ CONVERTIR null/undefined a string vacío de forma segura
          if (valor === null || valor === undefined) {
            valor = '';
          }

          reg[nombreColumna] = valor;
        }

        borradores.push(reg);
        console.log('  ✅ Borrador indexado exitosamente. Folio:', reg.FOLIO || 'S/F');
      }
    }

    console.log('📋 Total de borradores consolidados:', borradores.length);

    return {
      ok: true,
      borradores: borradores
    };

  } catch (err) {
    console.error('❌ Error en listarBorradores:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Elimina un borrador (libera el equipo)
 */
function eliminarBorrador(token, folio) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  var sheet = SHEETS.BITACORA();
  var datos = sheet.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === folio) {
      if (datos[i][7].toString() !== sesion.usuarioId) {
        return { ok: false, error: 'No tienes permiso para eliminar este borrador.' };
      }

      // Liberar la serie GPS
      var serieGPS = datos[i][4];
      if (serieGPS) {
        _liberarSerieGPS(serieGPS);
      }

      // Eliminar la fila
      sheet.deleteRow(i + 1);
      return { ok: true, mensaje: 'Borrador eliminado correctamente.' };
    }
  }
  return { ok: false, error: 'Borrador no encontrado.' };
}

/**
 * Convierte un borrador en reporte completo (cambia a "Listo para pago")
 */
function completarBorrador(token, folio, detalle) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  var sheet = SHEETS.BITACORA();
  var datos = sheet.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === folio) {
      if (datos[i][7].toString() !== sesion.usuarioId) {
        return { ok: false, error: 'No tienes permiso para modificar este borrador.' };
      }

      var fila = i + 1;

      // Actualizar detalle si se proporcionó
      if (detalle) {
        sheet.getRange(fila, 10).setValue(detalle);
      }

      // Cambiar estado a "Listo para pago" (ya tiene tipo y precio asignados)
      // Nota: El Revisor debe asignar el tipo de revisión
      sheet.getRange(fila, 13).setValue('Listo para pago');

      return { ok: true, mensaje: 'Borrador completado correctamente.' };
    }
  }
  return { ok: false, error: 'Borrador no encontrado.' };
}
// ============================================================
// 13. OCR - PROCESAMIENTO DE IMÁGENES
// ============================================================

/**
 * Procesa una imagen con OCR usando Google Drive
 * @param {string} base64Data - Imagen en base64
 * @param {string} nombreArchivo - Nombre del archivo
 * @returns {Object} { ok: boolean, texto: string, error: string }
 */
function procesarOCRImagen(base64Data, nombreArchivo) {
  try {
    console.log('=== INICIO procesarOCRImagen ===');
    console.log('Archivo:', nombreArchivo);

    // 1. Decodificar base64 a blob
    var base64Limpio = base64Data;
    if (base64Data.includes(',')) {
      base64Limpio = base64Data.split(',')[1];
    }

    var byteCharacters = Utilities.base64Decode(base64Limpio);
    var blob = Utilities.newBlob(byteCharacters, 'image/jpeg', nombreArchivo || 'imagen_ocr.jpg');

    // 2. Crear metadatos para Drive API v2
    var file = {
      title: 'OCR_Temporal_' + new Date().getTime(),
      mimeType: 'image/jpeg' // Forzamos formato de imagen compatible con OCR
    };

    console.log('Subiendo archivo a Drive con OCR...');

    // Insertar archivo y forzar OCR nativo
    var archivoOcr = Drive.Files.insert(file, blob, {
      ocr: true,
      ocrLanguage: 'es'
    });

    console.log('Archivo creado en Drive con ID:', archivoOcr.id);

    // 3. OBTENER EL TEXTO (Solución al error de permisos)
    // En lugar de DocumentApp, descargamos el recurso directamente como texto plano
    // Usamos el token de autenticación del propio script
    var url = "https://googleapis.com" + archivoOcr.id + "/export?mimeType=text/plain";
    var opciones = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };

    var respuesta = UrlFetchApp.fetch(url, opciones);
    var textoExtraido = respuesta.getContentText("UTF-8").trim();

    console.log('Texto extraído exitosamente.');

    // 4. Eliminar archivo temporal de la papelera directamente
    try {
      Drive.Files.remove(archivoOcr.id); // remove() lo borra permanentemente sin pasar por papelera
      console.log('Archivo temporal eliminado permanentemente');
    } catch (e) {
      console.warn('No se pudo eliminar archivo temporal:', e.message);
    }

    // 5. Extraer serie GPS (Tu función personalizada)
    var serieGPS = _extraerSerieGPS(textoExtraido);

    if (serieGPS) {
      return {
        ok: true,
        texto: serieGPS,
        textoCompleto: textoExtraido
      };
    } else if (textoExtraido.length > 0) {
      return {
        ok: false,
        error: 'No se detectó un código de serie GPS. El texto detectado fue: "' + textoExtraido.substring(0, 50) + '..."',
        textoCompleto: textoExtraido // Te lo devuelvo para que audites qué leyó
      };
    } else {
      return {
        ok: false,
        error: 'No se detectó texto en la imagen. Asegúrate de que la etiqueta esté bien enfocada.'
      };
    }

  } catch (error) {
    console.error('Error en OCR:', error);
    return {
      ok: false,
      error: 'Error al procesar OCR: ' + error.toString()
    };
  }
}

/**
 * Extrae una serie GPS de un texto o comentario de forma inteligente
 * Prioriza el formato estándar 4-3-3 (ej: ABCD-123-456 o abcd123456)
 * @param {string} texto - Texto o comentario que contiene la serie
 * @returns {string|null} Serie formateada en mayúsculas o null si no se halla
 */
function _extraerSerieGPS(texto) {
  if (!texto || texto.toString().trim() === '') return null;

  var textoStr = texto.toString().trim();

  // 1. PRIORIDAD MÁXIMA: Buscar el formato exacto de tu sistema 4-3-3 (con guiones)
  // Añadida la bandera 'i' para aceptar minúsculas
  var regexEstandar = /\b[A-Z0-9]{4}-[A-Z0-9]{3}-[A-Z0-9]{3}\b/gi;
  var coincidenciaEst = textoStr.match(regexEstandar);

  if (coincidenciaEst && coincidenciaEst.length > 0) {
    return coincidenciaEst[0].toUpperCase();
  }

  // 2. SEGUNDA PRIORIDAD: Buscar texto corrido de exactamente 10 caracteres alfanuméricos
  // 💡 NOTA: Exigimos que empiece con letras típicas de tus equipos (VG o CM) para EVITAR 
  // capturar palabras comunes como "VEHICULO" o "PROBLEMA".
  var regexCorrido = /\b(VG|CM)[A-Z0-9]{8}\b/gi;
  var coincidenciaCorrido = textoStr.match(regexCorrido);

  if (coincidenciaCorrido && coincidenciaCorrido.length > 0) {
    var codigo = coincidenciaCorrido[0].toUpperCase();
    // Lo formateamos automáticamente a tu estándar 4-3-3 para que haga match en el inventario
    return codigo.substring(0, 4) + '-' + codigo.substring(4, 7) + '-' + codigo.substring(7, 10);
  }

  // 3. TERCERA PRIORIDAD: Buscar variantes con guiones laxos (como tu paso 3 original)
  // Limitado por \b (límites de palabra) para no extraer pedazos de textos más largos
  var regexLaxo = /\b[A-Z0-9]{2,4}-[A-Z0-9]{2,4}-[A-Z0-9]{2,4}\b/gi;
  var coincidenciaLaxa = textoStr.match(regexLaxo);

  if (coincidenciaLaxa && coincidenciaLaxa.length > 0) {
    return coincidenciaLaxa[0].toUpperCase();
  }

  return null;
}
/**
 * Obtiene los parámetros de configuración para el frontend de forma segura
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: boolean, params: Object, error: string|null }
 */
function obtenerParametrosFrontend(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    // Intentamos leer la función interna de parámetros
    var params = typeof _leerParams === 'function' ? _leerParams() : null;
    if (!params) {
      return { ok: false, error: 'No se pudieron recuperar las variables de configuración.' };
    }

    // 🛠️ VALIDACIÓN DE BOOLEANOS ROBUSTA: Convierte a texto limpio y mayúsculas
    var compRaw = (params['COMPRESION_IMAGENES'] || '').toString().toUpperCase().trim();
    var activarCompresion = (params['COMPRESION_IMAGENES'] === true || compRaw === 'TRUE' || compRaw === 'SI' || compRaw === '1');

    // 🛠️ PROTECCIÓN CONTRA NaN: Aseguramos números reales con valores de respaldo
    var calidad = parseInt(params['CALIDAD_IMAGEN'], 10);
    var maxWidth = parseInt(params['MAX_WIDTH_IMAGEN'], 10);
    var maxHeight = parseInt(params['MAX_HEIGHT_IMAGEN'], 10);

    var frontendParams = {
      calidadImagen: isNaN(calidad) || calidad <= 0 ? 80 : Math.min(calidad, 100), // Rango seguro 1-100
      maxWidthImagen: isNaN(maxWidth) || maxWidth <= 0 ? 1200 : maxWidth,
      maxHeightImagen: isNaN(maxHeight) || maxHeight <= 0 ? 1200 : maxHeight,
      compresionImagenes: activarCompresion
    };

    return { ok: true, params: frontendParams };

  } catch (err) {
    console.error('Error crítico al obtener parámetros para frontend:', err);
    return { ok: false, error: 'Error al cargar configuraciones: ' + err.message };
  }
}

// ============================================================
// 14. VALIDACIÓN DE SERIE GPS EN TIEMPO REAL
// ============================================================

/**
 * Valida una serie GPS en tiempo real evaluando su último estado en el inventario
 * @param {string} token - Token de sesión
 * @param {string} serieGPS - Serie a validar (acepta con o sin guiones, mayúsculas o minúsculas)
 * @returns {Object} { ok: boolean, disponible: boolean, mensaje: string, estado: string, detalles: Object }
 */
function validarSerieGPSInventario(token, serieGPS) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (!serieGPS || serieGPS.toString().trim() === '') {
    return { ok: true, disponible: false, mensaje: 'Ingresa una serie GPS', estado: 'vacio' };
  }

  // 💡 MEJORA EN FORMATO: Limpiamos la serie para que la validación sea flexible y humana
  var serieBusquedaClean = serieGPS.toString().toUpperCase().replace(/-/g, '').trim();

  // Validamos que tenga exactamente 10 caracteres alfanuméricos (Estructura de tu formato 4-3-3)
  var regexEstructura = /^[A-Z0-9]{10}$/;
  if (!regexEstructura.test(serieBusquedaClean)) {
    return { ok: true, disponible: false, mensaje: 'Formato inválido. Debe tener 10 caracteres (ej: XXXX-XXX-XXX)', estado: 'formato_invalido' };
  }

  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }

    var datos = sheet.getDataRange().getValues();
    var ultimaCoincidencia = null;

    // Recorremos toda la hoja para capturar siempre el último movimiento (Compatibilidad con historiales)
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var serieCeldaRaw = fila;
      if (!serieCeldaRaw) continue;

      var serieCeldaClean = serieCeldaRaw.toString().toUpperCase().replace(/-/g, '').trim();

      if (serieCeldaClean === serieBusquedaClean) {
        // En lugar de hacer return inmediato, guardamos la fila. 
        // Si aparece más abajo, se actualizará con el estado más reciente.
        ultimaCoincidencia = {
          serieFormateadaBase: serieCeldaRaw.toString().toUpperCase().trim(),
          estadoRaw: fila || 'DESCONOCIDO',
          estadoClean: (fila || '').toString().toUpperCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""), // Remueve acentos
          economicoAsignado: (fila || '').toString().trim(),
          ticketGarantia: (fila || '').toString().trim()
        };
      }
    }

    // 📋 EVALUACIÓN FINAL DEL ÚLTIMO ESTADO ENCONTRADO
    if (ultimaCoincidencia) {
      var est = ultimaCoincidencia.estadoClean;
      var serieBonita = ultimaCoincidencia.serieFormateadaBase;
      var eco = ultimaCoincidencia.economicoAsignado;
      var tkt = ultimaCoincidencia.ticketGarantia;

      if (est === 'DISPONIBLE') {
        return {
          ok: true,
          disponible: true,
          mensaje: '✅ Serie disponible para instalar',
          estado: 'disponible',
          detalles: { estado: 'Disponible' }
        };
      }

      if (est === 'INSTALADO') {
        return {
          ok: true,
          disponible: false,
          mensaje: '⚠️ Serie ' + serieBonita + ' ya instalada en económico: ' + eco,
          estado: 'instalado',
          detalles: { estado: 'Instalado', economico: eco }
        };
      }

      if (est === 'GARANTIA') {
        return {
          ok: true,
          disponible: false,
          mensaje: '🚫 Serie ' + serieBonita + ' en garantía. Ticket: ' + (tkt || 'N/A'),
          estado: 'garantia',
          detalles: { estado: 'Garantía', ticket: tkt }
        };
      }

      if (est === 'BAJA') {
        return {
          ok: true,
          disponible: false,
          mensaje: '❌ Serie ' + serieBonita + ' dada de baja',
          estado: 'baja',
          detalles: { estado: 'Baja' }
        };
      }

      // Manejo de estados alterados o textos raros en la celda
      return {
        ok: true,
        disponible: false,
        mensaje: '⚠️ Estado no controlado para la serie ' + serieBonita + ': ' + ultimaCoincidencia.estadoRaw,
        estado: 'desconocido',
        detalles: { estado: ultimaCoincidencia.estadoRaw }
      };
    }

    // Si terminó el ciclo y no hay coincidencia
    return {
      ok: true,
      disponible: false,
      mensaje: '❌ Serie no encontrada en el inventario',
      estado: 'no_encontrada'
    };

  } catch (err) {
    console.error('Error al validar serie GPS:', err);
    return { ok: false, error: 'Error al validar: ' + err.message };
  }
}

// ============================================================
// 15. REEMPLAZO DE GPS
// ============================================================

/**
 * Procesa el reemplazo de un GPS con validaciones atómicas y alta velocidad
 * @param {string} token - Token de sesión
 * @param {Object} datos - { serieNueva, economico, serieVieja, motivo, esGarantia }
 * @returns {Object} Respuesta estandarizada de éxito o error
 */
function procesarReemplazoGPS(token, datos) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  if (sesion.rol > 3) return { ok: false, error: 'Sin permisos para realizar reemplazos.' };

  // Validación inicial de campos obligatorios
  if (!datos.serieNueva || !datos.economico || !datos.motivo) {
    return { ok: false, error: 'Faltan datos obligatorios (Serie nueva, económico o motivo).' };
  }

  try {
    var sheetInventario = SHEETS.INVENTARIO();
    var sheetBitacora = SHEETS.BITACORA();

    if (!sheetInventario || !sheetBitacora) {
      return { ok: false, error: 'No se encontraron las hojas necesarias (Inventario / Bitácora).' };
    }

    var inventario = sheetInventario.getDataRange().getValues();
    var bitacora = sheetBitacora.getDataRange().getValues();

    var filaNueva = -1;
    var filaVieja = -1;

    // 💡 HOMOLOGACIÓN SEGURA: Pasamos a mayúsculas y removemos todos los guiones
    var serieNuevaBusqueda = datos.serieNueva.toString().toUpperCase().replace(/-/g, '').trim();
    var serieViejaBusqueda = datos.serieVieja ? datos.serieVieja.toString().toUpperCase().replace(/-/g, '').trim() : null;
    var economicoBusqueda = datos.economico.toString().toUpperCase().trim();

    // 1. ESCANEO DEL INVENTARIO (Búsqueda simultánea en un solo bucle)
    for (var i = 1; i < inventario.length; i++) {
      var serieCeldaRaw = inventario[i][0];
      if (!serieCeldaRaw) continue;

      var serieCeldaClean = serieCeldaRaw.toString().toUpperCase().replace(/-/g, '').trim();

      // Buscar coincidencia de serie nueva
      if (serieCeldaClean === serieNuevaBusqueda) {
        filaNueva = i + 1;
      }
      // Buscar coincidencia de serie vieja (si se proporcionó una)
      if (serieViejaBusqueda && serieCeldaClean === serieViejaBusqueda) {
        filaVieja = i + 1;
      }
    }

    // 2. VALIDACIONES DE NEGOCIO ANTES DE ESCRIBIR
    if (filaNueva === -1) {
      return { ok: false, error: 'La serie nueva (' + datos.serieNueva + ') no existe en el inventario.' };
    }

    // Comprobamos disponibilidad de la nueva serie utilizando el array en memoria
    var estadoNueva = (inventario[filaNueva - 1][4] || '').toString().toUpperCase().trim();
    if (estadoNueva !== 'DISPONIBLE') {
      return { ok: false, error: 'La serie nueva (' + datos.serieNueva + ') no está disponible. Estado actual: ' + inventario[filaNueva - 1][4] };
    }

    // Si se mandó una serie vieja pero no se encontró en la hoja, frenamos el proceso para evitar inconsistencias
    if (serieViejaBusqueda && filaVieja === -1) {
      return { ok: false, error: 'La serie que deseas retirar (' + datos.serieVieja + ') no fue encontrada en el inventario.' };
    }

    var fechaActual = new Date();

    // 3. APLICAR CAMBIOS EN LA SERIE NUEVA (Escribimos de la Columna 5 a la 9 de un solo golpe)
    var valorColH_Nueva = inventario[filaNueva - 1][7]; // Resguardamos Columna H
    sheetInventario.getRange(filaNueva, 5, 1, 5).setValues([[
      'Instalado',         // Columna 5 (E) - Estado
      datos.economico,     // Columna 6 (F) - Económico asignado
      fechaActual,         // Columna 7 (G) - Fecha instalación
      valorColH_Nueva,     // Columna 8 (H) - Mantener intacto
      fechaActual          // Columna 9 (I) - Última actualización
    ]]);

    // 4. APLICAR CAMBIOS EN LA SERIE VIEJA (Si existe)
    if (filaVieja !== -1) {
      var obsActual = inventario[filaVieja - 1][10] || '';
      var nuevaObs = obsActual + ' | Reemplazado por ' + datos.serieNueva + ' - ' + datos.motivo;

      if (datos.esGarantia) {
        // Formato Garantía: Columna 5 a 11 (7 columnas de ancho)
        sheetInventario.getRange(filaVieja, 5, 1, 7).setValues([[
          'Garantía',                          // Columna 5 (E) - Estado
          inventario[filaVieja - 1][5] || '',  // Columna 6 (F) - Mantiene su económico por trazabilidad
          inventario[filaVieja - 1][6] || '',  // Columna 7 (G) - Mantiene su fecha de instalación
          'RPL-' + fechaActual.getTime(),      // Columna 8 (H) - Ticket de garantía único
          fechaActual,                         // Columna 9 (I) - Última actualización
          inventario[filaVieja - 1][9] || '',  // Columna 10 (J) - Columna intermedia intacta
          nuevaObs                             // Columna 11 (K) - Observaciones
        ]]);
      } else {
        // Formato Baja Común: Columna 5 a 11 (7 columnas de ancho)
        sheetInventario.getRange(filaVieja, 5, 1, 7).setValues([[
          'Baja',                              // Columna 5 (E) - Estado
          '',                                  // Columna 6 (F) - Liberar económico (vacío)
          '',                                  // Columna 7 (G) - Limpiar fecha de instalación (vacío)
          inventario[filaVieja - 1][7] || '',  // Columna 8 (H) - Mantener ticket actual si tiene
          fechaActual,                         // Columna 9 (I) - Última actualización
          inventario[filaVieja - 1][9] || '',  // Columna 10 (J) - Columna intermedia intacta
          nuevaObs                             // Columna 11 (K) - Observaciones
        ]]);
      }
    }

    // 5. ACTUALIZACIÓN EN BLOQUE DE LA BITÁCORA
    var filaBitacoraBorrador = -1;
    for (var k = 1; k < bitacora.length; k++) {
      var ecoBitacora = (bitacora[k][2] || '').toString().toUpperCase().trim();
      var estadoBitacora = (bitacora[k][12] || '').toString().trim();

      if (ecoBitacora === economicoBusqueda && estadoBitacora === 'Borrador') {
        filaBitacoraBorrador = k + 1;
        break; // Encontramos el folio en borrador, salimos del bucle
      }
    }

    if (filaBitacoraBorrador !== -1) {
      var detalleActual = bitacora[filaBitacoraBorrador - 1][9] || '';
      var notaReemplazo = '\n\n🔄 REEMPLAZO DE GPS:\n' +
        'Serie nueva: ' + datos.serieNueva + '\n' +
        'Serie retirada: ' + (datos.serieVieja || 'N/A') + '\n' +
        'Motivo: ' + datos.motivo + '\n' +
        (datos.esGarantia ? '📋 Marcado como GARANTÍA' : '❌ Dado de BAJA');

      // Columna 10 (J) es el índice 9 en base 0
      sheetBitacora.getRange(filaBitacoraBorrador, 10).setValue(detalleActual + notaReemplazo);
    }

    return {
      ok: true,
      mensaje: 'Reemplazo completado con éxito. ' + datos.serieNueva + ' instalada en económico ' + datos.economico,
      serieNueva: datos.serieNueva,
      serieVieja: datos.serieVieja || 'N/A',
      motivo: datos.motivo,
      esGarantia: datos.esGarantia
    };

  } catch (err) {
    console.error('Error crítico en procesarReemplazoGPS:', err);
    return { ok: false, error: 'Error interno del servidor al procesar el reemplazo: ' + err.message };
  }
}


// ============================================================
// FUNCIONES AUXILIARES
// ============================================================

/**
 * Obtiene todas las series GPS instaladas en un vehículo de forma segura
 * @param {string|number} economico - ID del vehículo
 * @returns {string[]} Array de series instaladas únicas
 */
function _obtenerTodasLasSeriesGPSporEconomico(economico) {
  var series = [];

  // Validation de seguridad inicial
  if (economico === undefined || economico === null || economico.toString().trim() === '') {
    return series;
  }

  var sheet = SHEETS.INVENTARIO();
  if (!sheet) return series;

  var datos = sheet.getDataRange().getValues();

  // 💡 HOMOLOGACIÓN CLAVE: Pasamos a mayúsculas para evitar fallas por minúsculas
  var economicoStr = economico.toString().toUpperCase().trim();

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var serieRaw = fila[0];
    if (serieRaw === undefined || serieRaw === null || serieRaw.toString().trim() === '') continue;

    var tipo = fila[1] || '';
    var estado = (fila[4] || '').toString().trim();
    var economicoAsignado = (fila[5] || '').toString().toUpperCase().trim();

    // Comparación homologada en mayúsculas y limpia de espacios
    if (economicoAsignado === economicoStr && estado === 'Instalado') {
      var tipoStr = tipo.toString().toUpperCase();

      // Filtramos únicamente los que correspondan a Gateways o VG
      if (tipoStr.indexOf('VG') !== -1 || tipoStr.indexOf('GATEWAY') !== -1) {
        var serieFormateada = serieRaw.toString().trim();

        // 💡 PROTECCIÓN CONTRA HISTORIALES: Evita duplicar la serie si aparece repetida en la hoja
        if (series.indexOf(serieFormateada) === -1) {
          series.push(serieFormateada);
        }
      }
    }
  }

  return series;
}


/**
 * Libera una serie GPS (cambia estado a Disponible y limpia el económico)
 * Busca la serie ignorando guiones para máxima compatibilidad
 */
function _liberarSerieGPS(serieGPS) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) {
    console.warn('❌ No se encontró la hoja INVENTARIO');
    return { ok: false, error: 'No se encontró la hoja INVENTARIO.' };
  }

  if (!serieGPS || serieGPS.toString().trim() === '') {
    console.warn('⚠️ Serie GPS vacía, no se puede liberar');
    return { ok: false, error: 'La serie GPS provista está vacía.' };
  }

  // 💡 LIMPIEZA: Quitamos guiones y espacios, convertimos a mayúsculas
  var serieBusquedaClean = serieGPS.toString().toUpperCase().replace(/[-_\s]/g, '').trim();

  console.log('🔍 Buscando serie para liberar:', serieBusquedaClean);

  var datos = sheet.getDataRange().getValues();
  var encontrado = false;
  var filaEncontrada = -1;

  for (var i = 1; i < datos.length; i++) {
    var serieCeldaRaw = datos[i][0];
    if (!serieCeldaRaw) continue;

    var serieCeldaClean = serieCeldaRaw.toString().toUpperCase().replace(/[-_\s]/g, '').trim();

    if (serieCeldaClean === serieBusquedaClean) {
      filaEncontrada = i + 1;
      encontrado = true;
      console.log('✅ Coincidencia encontrada en fila:', filaEncontrada);
      console.log('   Serie en inventario:', serieCeldaRaw);
      break;
    }
  }

  if (!encontrado) {
    console.warn('⚠️ NO se encontró la serie en inventario:', serieGPS);
    return { ok: false, error: 'La serie no fue hallada en el inventario.' };
  }

  // ✅ LIBERAR LA SERIE
  var fechaActual = new Date();
  var valorColumna8 = datos[filaEncontrada - 1][7] || '';

  // Escritura en bloque (Columnas 5 a 9)
  sheet.getRange(filaEncontrada, 5, 1, 5).setValues([[
    'Disponible',         // Columna 5 - ESTADO
    '',                   // Columna 6 - ECONOMICO_ASIGNADO (se limpia)
    '',                   // Columna 7 - FECHA_INSTALACION (se limpia)
    valorColumna8,        // Columna 8 - TICKET_GARANTIA (se mantiene)
    fechaActual           // Columna 9 - ULTIMA_ACTUALIZACION
  ]]);

  console.log('✅ Serie LIBERADA exitosamente:', serieGPS);
  console.log('   Económico liberado');

  return { ok: true, error: null };
}


/**
 * Actualiza el estado del GPS en Inventario al ser instalado
 * Busca la serie ignorando guiones para máxima compatibilidad
 */
function _actualizarEstadoGPS(serieGPS, economico) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) {
    console.warn('❌ No se encontró la hoja INVENTARIO');
    return { ok: false, error: 'No se encontró la hoja INVENTARIO.' };
  }

  if (!serieGPS || serieGPS.toString().trim() === '') {
    console.warn('⚠️ Serie GPS vacía, no se puede actualizar');
    return { ok: false, error: 'La serie GPS provista está vacía.' };
  }

  // 💡 LIMPIEZA: Quitamos guiones y espacios, convertimos a mayúsculas
  var serieBusquedaClean = serieGPS.toString().toUpperCase().replace(/[-_\s]/g, '').trim();
  var economicoStr = (economico || '').toString().trim();

  console.log('🔍 Buscando serie limpia:', serieBusquedaClean);
  console.log('📌 Para económico:', economicoStr);

  var datos = sheet.getDataRange().getValues();
  var encontrado = false;
  var filaEncontrada = -1;

  for (var i = 1; i < datos.length; i++) {
    var serieCeldaRaw = datos[i][0];
    if (!serieCeldaRaw) continue;

    // 💡 LIMPIAMOS LA SERIE DEL INVENTARIO DE LA MISMA FORMA
    var serieCeldaClean = serieCeldaRaw.toString().toUpperCase().replace(/[-_\s]/g, '').trim();

    // Comparación sin guiones
    if (serieCeldaClean === serieBusquedaClean) {
      filaEncontrada = i + 1;
      encontrado = true;
      console.log('✅ Coincidencia encontrada en fila:', filaEncontrada);
      console.log('   Serie en inventario:', serieCeldaRaw);
      break;
    }
  }

  if (!encontrado) {
    console.warn('⚠️ NO se encontró la serie en inventario:', serieGPS);
    console.warn('   Buscando como:', serieBusquedaClean);
    console.warn('   Revisa que la serie esté registrada en 📦_Inventario_GPS');
    return { ok: false, error: 'La serie no fue hallada en el inventario.' };
  }

  // ✅ ACTUALIZAR LA SERIE ENCONTRADA
  var fechaActual = new Date();
  var valorColumna8 = datos[filaEncontrada - 1][7] || '';

  // Escritura en bloque de la fila (Columnas 5 a 9)
  sheet.getRange(filaEncontrada, 5, 1, 5).setValues([[
    'Instalado',          // Columna 5 - ESTADO
    economicoStr,         // Columna 6 - ECONOMICO_ASIGNADO
    fechaActual,          // Columna 7 - FECHA_INSTALACION
    valorColumna8,        // Columna 8 - TICKET_GARANTIA (se mantiene)
    fechaActual           // Columna 9 - ULTIMA_ACTUALIZACION
  ]]);

  console.log('✅ Serie INSTALADA exitosamente:', serieGPS);
  console.log('   Asignada a económico:', economicoStr);

  return { ok: true, error: null };
}

/**
 * Verifica la disponibilidad de una serie GPS de forma robusta basándose en su último estado
 * @param {string} serieGPS - Código de serie a verificar
 * @returns {Object} { disponible: boolean, mensaje: string }
 */
function _verificarDisponibilidadSerie(serieGPS) {
  var sheet = SHEETS.INVENTARIO();
  if (!sheet) {
    return { disponible: false, mensaje: 'No se encontró la hoja de inventario.' };
  }

  if (!serieGPS || serieGPS.toString().trim() === '') {
    return { disponible: false, mensaje: 'La serie GPS provista está vacía.' };
  }

  // Homologación de la serie de búsqueda (sin guiones)
  var serieBusquedaClean = serieGPS.toString().toUpperCase().replace(/-/g, '').trim();
  var datos = sheet.getDataRange().getValues();

  // Variables para almacenar la última coincidencia encontrada
  var ultimaCoincidencia = null;

  // Recorremos toda la hoja para asegurarnos de capturar el estado más reciente (al final de la hoja)
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var serieCeldaRaw = fila[0];
    if (serieCeldaRaw === undefined || serieCeldaRaw === null) continue;

    var serieCeldaClean = serieCeldaRaw.toString().toUpperCase().replace(/-/g, '').trim();

    if (serieCeldaClean === serieBusquedaClean) {
      // En lugar de salir con un return, guardamos los datos de esta fila. 
      // Si la serie vuelve a aparecer más abajo, esta variable se sobrescribirá con el estado más nuevo.
      ultimaCoincidencia = {
        serieFormateadaBase: serieCeldaRaw.toString().toUpperCase().trim(),
        estadoRaw: fila[4] || '',
        estado: (fila[4] || '').toString().toUpperCase().trim(),
        economicoAsignado: (fila[5] || 'N/A').toString().trim()
      };
    }
  }

  // 📋 EVALUACIÓN FINAL: Si encontramos la serie en alguna parte de la hoja
  if (ultimaCoincidencia) {
    var est = ultimaCoincidencia.estado;
    var serieBonita = ultimaCoincidencia.serieFormateadaBase;

    if (est === 'DISPONIBLE') {
      return { disponible: true, mensaje: 'Serie disponible' };
    }
    if (est === 'INSTALADO') {
      return { disponible: false, mensaje: 'La serie ' + serieBonita + ' ya está instalada en el económico ' + ultimaCoincidencia.economicoAsignado };
    }
    if (est === 'GARANTÍA' || est === 'GARANTIA') {
      return { disponible: false, mensaje: 'La serie ' + serieBonita + ' está en garantía' };
    }
    if (est === 'BAJA') {
      return { disponible: false, mensaje: 'La serie ' + serieBonita + ' está dada de baja' };
    }

    // Estado no controlado
    return { disponible: false, mensaje: 'La serie ' + serieBonita + ' tiene un estado no controlado: ' + ultimaCoincidencia.estadoRaw };
  }

  // Si el ciclo terminó y la variable sigue vacía, el equipo nunca ha sido registrado
  return { disponible: false, mensaje: 'La serie ' + serieGPS.toUpperCase() + ' no existe en el inventario' };
}


/**
 * Valida formato de serie GPS (4-3-3) ej: ABCD-123-456
 * @param {string|number} serie - Serie del dispositivo a validar
 * @returns {boolean} True si cumple el formato, False en caso contrario
 */
function _validarSerieGPS(serie) {
  // 1. Protección contra valores nulos o indefinidos
  if (serie === undefined || serie === null) return false;

  // 2. Limpieza y estandarización del texto
  // Convierte a texto, elimina espacios en los extremos y pasa todo a mayúsculas
  var serieLimpia = serie.toString().trim().toUpperCase();

  // 3. Expresión regular (El patrón se mantiene idéntico a tu lógica original)
  var regex = /^[A-Z0-9]{4}-[A-Z0-9]{3}-[A-Z0-9]{3}$/;

  return regex.test(serieLimpia);
}


/**
 * Verifica si una serie está bloqueada por garantía
 * @param {string|number} serieGPS - Serie del dispositivo a buscar
 * @returns {boolean} True si está bloqueada por garantía, False en caso contrario
 */
function _estaBloqueadaPorGarantia(serieGPS) {
  // 1. Validación de seguridad para evitar errores con valores nulos o vacíos
  if (serieGPS === undefined || serieGPS === null || serieGPS.toString().trim() === '') {
    return false;
  }

  var sheet = SHEETS.INVENTARIO();
  if (!sheet) return false;

  var datos = sheet.getDataRange().getValues();
  var serieBusqueda = serieGPS.toString().toUpperCase().trim();

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var serieExistente = fila[0];

    if (serieExistente === undefined || serieExistente === null) continue;

    // Normalizamos la serie de la hoja para una comparación exacta
    var serieExistenteStr = serieExistente.toString().toUpperCase().trim();

    if (serieExistenteStr === serieBusqueda) {
      var ticket = fila[7];
      var estado = (fila[4] || '').toString().trim();

      // ✅ REGLA DE NEGOCIO: Si tiene ticket y el estado es 'Garantía'
      if (ticket && ticket.toString().trim() !== '' && estado === 'Garantía') {
        return true;
      }

      // ❌ ELIMINADO EL 'return false;' DE AQUÍ. 
      // Si la serie coincide pero no está en garantía, permitimos que el ciclo siga buscando 
      // por si existe otro registro más adelante en la hoja.
    }
  }

  // Si terminó de revisar TODA la hoja y ninguna fila cumplió con la regla, entonces no está bloqueada.
  return false;
}


/**
 * Convierte las imágenes a formato JPEG para estandarizar su peso
 * @param {Blob[]} archivos - Matriz de Blobs de imágenes
 * @param {number} calidad - Parámetro estético (no altera bytes nativamente)
 * @returns {Blob[]} Matriz de archivos convertidos
 */
function _comprimirImagenes(archivos, calidad) {
  if (!archivos || !Array.isArray(archivos)) return [];

  return archivos.map(function (archivo) {
    try {
      var tipo = archivo.getContentType() || '';
      // Si ya es una imagen, la forzamos a convertirse en JPEG (suele pesar menos que PNG)
      if (tipo.indexOf('image/') !== -1 && tipo !== 'image/jpeg') {
        return archivo.getAs('image/jpeg');
      }
      return archivo;
    } catch (e) {
      console.error('Error al procesar archivo en compresión:', e);
      return archivo; // Si falla, devuelve el original para no romper el flujo
    }
  });
}


/**
 * Obtiene los tipos de equipo activos desde 📋_Tipos_Equipo
 * @param {string} token - Token de sesión
 * @param {string} categoria - Categoría a filtrar (opcional)
 * @returns {Object} { ok: true, tipos: [...] }
 */
function obtenerTiposEquipo(token, categoria) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var sheet = SHEETS.TIPOS_EQUIPO();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja 📋_Tipos_Equipo.' };
    }

    var datos = sheet.getDataRange().getValues();
    var tipos = [];

    // Convertimos la categoría a mayúsculas una sola vez
    var categoriaUpper = categoria ? categoria.toString().toUpperCase().trim() : null;

    for (var i = 1; i < datos.length; i++) {
      var clave = datos[i][0];
      var nombre = datos[i][1];
      var descripcion = datos[i][2];
      var activoRaw = datos[i][3];

      // 🛠️ VALIDACIÓN ROBUSTA DE BOOLEANOS: Acepta checkboxes reales, texto 'true', 'si' o 'activo'
      var activoStr = (activoRaw || '').toString().toUpperCase().trim();
      var isActivo = (activoRaw === true || activoStr === 'TRUE' || activoStr === 'SI' || activoStr === 'ACTIVO');

      // Solo tipos activos y con clave válida
      if (clave && isActivo) {

        var claveStr = clave.toString().toUpperCase().trim();
        var nombreStr = (nombre || '').toString().toUpperCase().trim();

        // Si no hay filtro, se incluye por defecto. Si hay filtro, empieza en false hasta que se demuestre lo contrario.
        var incluir = !categoriaUpper;

        // ✅ FILTRAR POR PATRONES (Uso de un switch para mayor claridad y velocidad)
        if (categoriaUpper) {
          switch (categoriaUpper) {
            case 'GATEWAY':
              if (claveStr.indexOf('VG') === 0 || nombreStr.indexOf('GATEWAY') !== -1) incluir = true;
              break;
            case 'CAMARA':
              if (claveStr.indexOf('CM') === 0 || nombreStr.indexOf('CAMARA') !== -1) incluir = true;
              break;
            case 'ARNES':
              if (claveStr.indexOf('ARNES') !== -1 || nombreStr.indexOf('ARNES') !== -1) incluir = true;
              break;
            case 'BOTON':
              if (claveStr.indexOf('BTN') !== -1 || nombreStr.indexOf('BOTON') !== -1) incluir = true;
              break;
            case 'CORTE':
              if (claveStr.indexOf('EI') === 0 || nombreStr.indexOf('CORTE') !== -1) incluir = true;
              break;
            default:
              // Si la categoría existe pero no coincide con ninguna regla conocida, se queda en false.
              incluir = false;
              break;
          }
        }

        // 💡 UNIFICACIÓN DE CARGA
        if (incluir) {
          tipos.push({
            clave: clave.toString().trim(),
            nombre: (nombre || '').toString().trim(),
            descripcion: (descripcion || '').toString().trim()
          });
        }
      }
    }

    return { ok: true, tipos: tipos };

  } catch (err) {
    console.error('Error al obtener tipos de equipo:', err);
    return { ok: false, error: err.message };
  }
}


/** Pendiente de revisar
 * Obtiene los tipos de equipo filtrados por categoría de forma segura
 */
function obtenerTiposEquipoPorCategoria(token, categoria) {
  var tiposResp = obtenerTiposEquipo(token);
  if (!tiposResp.ok) return tiposResp;

  var categoriaUpper = (categoria || '').toString().toUpperCase().trim();

  var filtrados = tiposResp.tipos.filter(function (t) {
    // 💡 SOLUCIÓN: Convertimos la clave a String para evitar el Crash de .startsWith() en celdas numéricas
    var claveStr = (t.clave || '').toString().toUpperCase();
    return claveStr.indexOf(categoriaUpper) === 0; // Equivalente seguro y rápido a startsWith
  });

  return { ok: true, tipos: filtrados };
}
/**
 * Obtiene todos los dispositivos instalados en un vehículo (Versión Optimizada)
 * @param {string} token - Token de sesión
 * @param {string} economico - ID del vehículo (ej. G-361)
 * @returns {Object} { ok: boolean, gateway, camara, accesorios }
 */
function obtenerDispositivosPorEconomico(token, economico) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    if (!economico) return { ok: false, error: 'El número económico es requerido.' };

    var resultado = {
      ok: true,
      gateway: null,
      camara: null,
      accesorios: {} // Almacena el tipo de accesorio y si está presente
    };

    // Homologamos el parámetro de búsqueda a mayúsculas y sin espacios
    var economicoStr = economico.toString().toUpperCase().trim();
    console.log('🔍 Buscando dispositivos para económico homologado:', economicoStr);

    // ============================================================
    // 1. BUSCAR EN INVENTARIO GPS (Gateway y Cámara)
    // ============================================================
    var sheetInventario = SHEETS.INVENTARIO();
    if (sheetInventario) {
      var datosInventario = sheetInventario.getDataRange().getValues();

      for (var i = 1; i < datosInventario.length; i++) {
        var serie = datosInventario[i][0];
        var tipo = datosInventario[i][1] || '';
        var estado = datosInventario[i][4] || '';
        var economicoAsignado = datosInventario[i][5] || '';

        // 💡 OPTIMIZACIÓN: Comparación tolerante a mayúsculas/minúsculas y espacios
        var economicoAsignadoStr = economicoAsignado.toString().toUpperCase().trim();
        var estadoStr = estado.toString().toUpperCase().trim();

        if (economicoAsignadoStr === economicoStr && estadoStr === 'INSTALADO') {
          var tipoStr = tipo.toString().toUpperCase();

          if (tipoStr.indexOf('GATEWAY') !== -1 || tipoStr.indexOf('VG') !== -1) {
            resultado.gateway = {
              serie: (serie || '').toString(),
              tipo: tipo.toString()
            };
            console.log('✅ Gateway encontrado:', resultado.gateway);
          } else if (tipoStr.indexOf('CAMARA') !== -1 || tipoStr.indexOf('CM') !== -1) {
            resultado.camara = {
              serie: (serie || '').toString(),
              tipo: tipo.toString()
            };
            console.log('✅ Cámara encontrada:', resultado.camara);
          }
        }
      }
    }

    // ============================================================
    // 2. BUSCAR EN ACCESORIOS STOCK (Listas por comas tolerantes)
    // ============================================================
    var sheetAccesorios = SHEETS.ACCESORIOS();
    if (sheetAccesorios) {
      var datosAccesorios = sheetAccesorios.getDataRange().getValues();

      for (var j = 1; j < datosAccesorios.length; j++) {
        var accesorio = datosAccesorios[j][0];
        var tipo = datosAccesorios[j][1] || '';
        var economicoAsignado = datosAccesorios[j][5] || '';

        var economicoAsignadoStr = economicoAsignado.toString().toUpperCase().trim();

        if (economicoAsignadoStr) {
          // Dividimos la lista por comas y limpiamos espacios de cada elemento en mayúsculas
          var listaEconomicos = economicoAsignadoStr.split(',').map(function (e) {
            return e.trim();
          });

          // Si el vehículo actual se encuentra mapeado en el array de asignaciones
          if (listaEconomicos.indexOf(economicoStr) !== -1) {
            var nombreTipo = tipo.toString();
            // Registramos el accesorio de forma segura
            resultado.accesorios[nombreTipo] = true;
            console.log('✅ Accesorio asociado encontrado:', nombreTipo);
          }
        }
      }
    }

    console.log('📋 Resultado consolidado enviado al cliente:', resultado);
    return resultado;

  } catch (err) {
    console.error('Error en obtenerDispositivosPorEconomico:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Actualiza el stock de accesorios en base a instalaciones o desinstalaciones
 * @param {string} tipoAccesorio - Tipo de accesorio a buscar
 * @param {string|number} economico - ID del vehículo
 * @param {string} tipo - 'instalacion' o 'desinstalacion'
 * @returns {Object} { ok: boolean, error: string|null }
 */
function _actualizarStockAccesorios(tipoAccesorio, economico, tipo) {
  // Validación de parámetros iniciales
  if (!tipoAccesorio || !economico || !tipo) {
    return { ok: false, error: 'Parámetros insuficientes para actualizar accesorios.' };
  }

  var sheet = SHEETS.ACCESORIOS();
  if (!sheet) {
    console.warn('❌ No se encontró la hoja de accesorios');
    return { ok: false, error: 'No se encontró la hoja de accesorios' };
  }

  var datos = sheet.getDataRange().getValues();
  var economicoStr = economico.toString().trim();
  var tipoAccesorioStr = tipoAccesorio.toString().toUpperCase().trim();

  console.log('🔧 _actualizarStockAccesorios - EJECUTADA');
  console.log('📌 tipoAccesorio:', tipoAccesorioStr);
  console.log('📌 economico:', economicoStr);
  console.log('📌 tipo:', tipo);

  for (var i = 1; i < datos.length; i++) {
    var accesorioExistente = datos[i][0];
    if (!accesorioExistente) continue;

    var accesorioExistenteStr = accesorioExistente.toString().toUpperCase().trim();

    if (accesorioExistenteStr === tipoAccesorioStr) {
      var fila = i + 1;

      // Leer valores actuales
      var stockTotal = Number(datos[i][2]) || 0;      // Columna C
      var asignados = Number(datos[i][3]) || 0;       // Columna D
      var disponibles = Number(datos[i][4]) || 0;     // Columna E
      var economicoAsignado = datos[i][5] || '';      // Columna F

      console.log('✅ Accesorio encontrado - Fila:', fila);

      var nuevosAsignados = asignados;
      var nuevosDisponibles = disponibles;
      var listaEconomicos = economicoAsignado ? economicoAsignado.split(',') : [];
      var listaEconomicosLimpios = listaEconomicos.map(function (e) { return e.trim(); });
      var nuevaLista = [];

      if (tipo === 'instalacion') {
        nuevosAsignados = asignados + 1;
        nuevosDisponibles = disponibles - 1;

        // Validaciones críticas de negocio
        if (nuevosAsignados > stockTotal || nuevosDisponibles < 0) {
          console.error('❌ No hay suficiente stock disponible.');
          return { ok: false, error: 'Stock insuficiente para el accesorio: ' + tipoAccesorioStr };
        }

        // Agregar económico si no existe
        if (listaEconomicosLimpios.indexOf(economicoStr) === -1) {
          listaEconomicosLimpios.push(economicoStr);
        }
        nuevaLista = listaEconomicosLimpios;

      } else if (tipo === 'desinstalacion') {
        nuevosAsignados = Math.max(0, asignados - 1);
        nuevosDisponibles = disponibles + 1;

        // Quitar económico de la lista
        nuevaLista = listaEconomicosLimpios.filter(function (e) {
          return e !== economicoStr;
        });
      } else {
        return { ok: false, error: 'Tipo de operación no válido.' };
      }

      // ⚡ OPTIMIZACIÓN CRÍTICA: Escribir toda la fila de un solo golpe (Columnas D a G)
      // Evitamos hacer 4 llamadas pesadas de .setValue() y hacemos 1 sola operación por bloques.
      var filaActualizacion = [
        nuevosAsignados,          // Columna D (Asignados)
        nuevosDisponibles,         // Columna E (Disponibles)
        nuevaLista.join(', '),     // Columna F (Lista Económicos)
        new Date()                 // Columna G (Fecha Actualización)
      ];

      // getRange(fila, columnaInicial, filas, columnas) -> Fila actual, columna 4 (D), 1 fila de alto, 4 columnas de ancho
      sheet.getRange(fila, 4, 1, 4).setValues([filaActualizacion]);

      console.log('✅ Inventario actualizado con éxito en fila:', fila);
      return { ok: true, error: null };
    }
  }

  console.warn('⚠️ No se encontró el accesorio:', tipoAccesorioStr);
  return { ok: false, error: 'El accesorio ' + tipoAccesorioStr + ' no existe en el catálogo.' };
}
/**
 * Obtiene el tipo de equipo de una serie desde el inventario
 * @param {string} token - Token de sesión
 * @param {string} serie - Serie del dispositivo
 * @returns {Object} { ok, existe, tipo, estado, modelo, economico, mensaje, error }
 */
function obtenerTipoPorSerie(token, serie) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Corrección de consistencia: Retorna los campos por defecto vacíos en lugar de omitirlos
  if (!serie || serie.trim() === '') {
    return { ok: true, existe: false, tipo: null, estado: null, modelo: null, economico: null, mensaje: 'Serie vacía' };
  }

  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }

    var datos = sheet.getDataRange().getValues();
    var serieBusqueda = serie.toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var serieExistente = datos[i][0];
      if (serieExistente === undefined || serieExistente === null) continue;

      if (serieExistente.toString().toUpperCase().trim() === serieBusqueda) {
        var tipo = datos[i][1] || '';
        var modelo = datos[i][2] || ''; // Asumiendo columna 3 según tu código original
        var estado = datos[i][4] || ''; // Asumiendo columna 5 según tu código original
        var economico = datos[i][5] || ''; // Asumiendo columna 6 según tu código original

        return {
          ok: true,
          existe: true,
          tipo: tipo.toString().trim(),
          estado: estado.toString().trim(),
          modelo: modelo.toString().trim(),
          economico: economico.toString().trim()
        };
      }
    }

    // Corrección de consistencia: Retorna la estructura completa de campos aunque no exista
    return {
      ok: true,
      existe: false,
      tipo: null,
      estado: null,
      modelo: null,
      economico: null,
      mensaje: 'La serie ' + serie + ' no existe en el inventario.'
    };

  } catch (err) {
    console.error('Error al obtener tipo por serie:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Verifica si un vehículo ya tiene equipos instalados
 * @param {string|number} economico - ID del vehículo
 * @returns {Object} { tieneEquipos: boolean, gateway: string, camara: string }
 */
function _verificarEquiposVehiculo(economico) {
  var resultado = {
    tieneEquipos: false,
    gateway: null,
    camara: null
  };

  // 1. Validación de seguridad
  if (economico === undefined || economico === null) return resultado;

  var sheet = SHEETS.INVENTARIO();
  if (!sheet) return resultado;

  // 2. Obtención y normalización de datos
  var datos = sheet.getDataRange().getValues();
  var economicoStr = economico.toString().trim().toUpperCase(); // Normalizado a mayúsculas

  for (var i = 1; i < datos.length; i++) {
    // Evita errores si la celda de económico asignado está vacía
    var economicoAsignado = datos[i][5];
    if (!economicoAsignado) continue;

    var economicoAsignadoStr = economicoAsignado.toString().trim().toUpperCase();
    var estado = (datos[i][4] || '').toString().trim();

    // 3. Validación de coincidencia
    if (economicoAsignadoStr === economicoStr && estado === 'Instalado') {
      resultado.tieneEquipos = true;

      var serie = (datos[i][0] || '').toString().trim();
      var tipoStr = (datos[i][1] || '').toString().toUpperCase();

      // Identificación de equipos
      if (tipoStr.indexOf('VG') !== -1 || tipoStr.indexOf('GATEWAY') !== -1) {
        resultado.gateway = serie;
      } else if (tipoStr.indexOf('CM') !== -1 || tipoStr.indexOf('CAMARA') !== -1) {
        resultado.camara = serie;
      }

      // OPTIMIZACIÓN: Si ya encontramos ambos equipos, rompemos el ciclo inmediatamente
      if (resultado.gateway && resultado.camara) {
        break;
      }
    }
  }

  return resultado;
}
/**
 * Obtiene TODOS los dispositivos (Gateway y Cámara) instalados en un vehículo
 * @param {string} economico - ID del vehículo
 * @returns {Object} { gateway: string|null, camara: string|null }
 */
function _obtenerTodosLosDispositivosPorEconomico(economico) {
  var resultado = {
    gateway: null,
    camara: null
  };

  var sheet = SHEETS.INVENTARIO();
  if (!sheet) return resultado;

  var datos = sheet.getDataRange().getValues();
  var economicoStr = economico.toString().toUpperCase().trim();

  for (var i = 1; i < datos.length; i++) {
    var serie = datos[i][0];
    if (!serie) continue;

    var tipo = (datos[i][1] || '').toString().toUpperCase();
    var estado = (datos[i][4] || '').toString().toUpperCase().trim();
    var economicoAsignado = (datos[i][5] || '').toString().toUpperCase().trim();

    if (economicoAsignado === economicoStr && estado === 'INSTALADO') {
      var serieStr = serie.toString().trim();

      if (tipo.indexOf('VG') !== -1 || tipo.indexOf('GATEWAY') !== -1) {
        resultado.gateway = serieStr;
      } else if (tipo.indexOf('CM') !== -1 || tipo.indexOf('CAMARA') !== -1) {
        resultado.camara = serieStr;
      }
    }
  }

  return resultado;
}
/**
 * Obtiene las URLs de las imágenes de una carpeta de Drive
 * @param {string} token - Token de sesión
 * @param {string} folderUrl - URL de la carpeta en Drive
 * @returns {Object} { ok, fotos: [{nombre, url, tipo, thumbnail}] }
 */
function obtenerFotosCarpeta(token, folderUrl) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (!folderUrl || folderUrl.trim() === '') {
    return { ok: true, fotos: [] };
  }

  try {
    // Extraer el ID de la carpeta
    var match = folderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (!match || !match[1]) {
      return { ok: true, fotos: [] };
    }

    var folderId = match[1];
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var fotos = [];

    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();

      // Solo imágenes
      if (mimeType.startsWith('image/')) {
        // ✅ OBTENER URL DIRECTA DE LA IMAGEN
        var fileId = file.getId();
        var urlDirecta = 'https://drive.google.com/uc?export=view&id=' + fileId;
        // También obtener la URL de la miniatura (mejor rendimiento)
        var thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w200';

        fotos.push({
          nombre: file.getName(),
          url: file.getUrl(),
          urlDirecta: urlDirecta,
          thumbnail: thumbnailUrl,
          tipo: mimeType,
          id: fileId
        });
      }
    }

    return { ok: true, fotos: fotos };

  } catch (err) {
    console.error('Error al obtener fotos de carpeta:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Verifica si un vehículo tiene equipos instalados (Gateway o Cámara)
 * @param {string} token - Token de sesión
 * @param {string} economico - ID del vehículo
 * @returns {Object} { tieneEquipos: boolean, gateway: string|null, camara: string|null }
 */
function verificarEquiposVehiculoBackend(token, economico) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (!economico || economico.trim() === '') {
    return { ok: true, tieneEquipos: false, gateway: null, camara: null };
  }

  try {
    var resultado = {
      ok: true,
      tieneEquipos: false,
      gateway: null,
      camara: null
    };

    var sheet = SHEETS.INVENTARIO();
    if (!sheet) return resultado;

    var datos = sheet.getDataRange().getValues();
    var economicoStr = economico.toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var serie = datos[i][0];
      if (!serie) continue;

      var tipo = (datos[i][1] || '').toString().toUpperCase();
      var estado = (datos[i][4] || '').toString().toUpperCase().trim();
      var economicoAsignado = (datos[i][5] || '').toString().toUpperCase().trim();

      if (economicoAsignado === economicoStr && estado === 'INSTALADO') {
        resultado.tieneEquipos = true;
        var serieStr = serie.toString().trim();

        if (tipo.indexOf('VG') !== -1 || tipo.indexOf('GATEWAY') !== -1) {
          resultado.gateway = serieStr;
        } else if (tipo.indexOf('CM') !== -1 || tipo.indexOf('CAMARA') !== -1) {
          resultado.camara = serieStr;
        }
      }
    }

    return resultado;

  } catch (err) {
    console.error('Error en verificarEquiposVehiculoBackend:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// INVENTARIO GPS - BACKEND
// ============================================================

/**
 * Obtiene todo el inventario GPS
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, equipos: [...] }
 */
/**
 * Obtiene todo el inventario
 */
function obtenerInventarioGPS(token) {
  console.log('📦 obtenerInventarioGPS - INICIO');

  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var sheetInventario = SHEETS.INVENTARIO();
    if (!sheetInventario) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }

    // ✅ LEER TIPO_UNIDAD DESDE 📋_Catalogo_Vehiculos (por vehículo)
    var catalogoSheet = SS.getSheetByName('📋_Catalogo_Vehiculos');
    var tiposUnidad = {};
    var vehiculosActivos = {};

    if (catalogoSheet) {
      var catalogoData = catalogoSheet.getDataRange().getValues();
      // Saltar encabezado (fila 1)
      for (var c = 1; c < catalogoData.length; c++) {
        var eco = (catalogoData[c][0] || '').toString().toUpperCase().trim();
        var tipoUnidad = (catalogoData[c][10] || '').toString().trim(); // Columna K = TIPO_UNIDAD
        var estado = (catalogoData[c][8] || '').toString().toUpperCase().trim(); // Columna I = ESTADO
        if (eco) {
          tiposUnidad[eco] = tipoUnidad;
          vehiculosActivos[eco] = (estado === 'ACTIVO');
        }
      }
      console.log('📌 Tipos de unidad desde catálogo:', Object.keys(tiposUnidad).length);
    }

    var datos = sheetInventario.getDataRange().getValues();
    var equipos = [];

    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (!f[0]) continue;

      var economico = (f[5] || '').toString().trim();
      var economicoUpper = economico.toUpperCase().trim();

      // ✅ Obtener el tipo de unidad del vehículo desde el catálogo
      // Si el vehículo existe en el catálogo, usar su tipo, si no, vacío
      var tipoUnidad = '';
      if (economico && tiposUnidad[economicoUpper] !== undefined) {
        tipoUnidad = tiposUnidad[economicoUpper];
      }

      equipos.push({
        serie: f[0] || '',
        tipo: f[1] || '',
        modelo: f[2] || '',
        imei: f[3] || '',
        estado: f[4] || '',
        economico: economico,
        fechaInstalacion: f[6] ? Utilities.formatDate(new Date(f[6]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        ticketGarantia: f[7] || '',
        fechaGarantia: f[8] ? Utilities.formatDate(new Date(f[8]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        ultimaActualizacion: f[9] ? Utilities.formatDate(new Date(f[9]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        observaciones: f[10] || '',
        // ✅ TIPO_UNIDAD desde el catálogo (por vehículo)
        tipoUnidad: tipoUnidad
      });
    }

    console.log('📌 Equipos procesados:', equipos.length);
    return { ok: true, equipos: equipos };

  } catch (err) {
    console.error('Error en obtenerInventarioGPS:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Obtiene un equipo específico del inventario
 * @param {string} token - Token de sesión
 * @param {string} serie - Serie del equipo
 * @returns {Object} { ok: true, equipo: {...} }
 */
/**
 * Obtiene un equipo por serie
 */
function obtenerEquipoInventario(token, serie) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) return { ok: false, error: 'No se encontró la hoja de inventario.' };

    var datos = sheet.getDataRange().getValues();
    var busqueda = serie.toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (!f[0]) continue;
      if (f[0].toString().toUpperCase().trim() === busqueda) {
        return {
          ok: true,
          equipo: {
            serie: f[0] || '',
            tipo: f[1] || '',
            modelo: f[2] || '',
            imei: f[3] || '',
            estado: f[4] || '',
            economico: f[5] || '',
            fechaInstalacion: f[6] ? Utilities.formatDate(new Date(f[6]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
            ticketGarantia: f[7] || '',
            fechaGarantia: f[8] ? Utilities.formatDate(new Date(f[8]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
            ultimaActualizacion: f[9] ? Utilities.formatDate(new Date(f[9]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
            observaciones: f[10] || '',
            tipoUnidad: f[11] || ''
          }
        };
      }
    }
    return { ok: false, error: 'Equipo no encontrado.' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Actualiza el estado de un equipo en el inventario (solo Admin/Revisor)
 * @param {string} token - Token de sesión
 * @param {string} serie - Serie del equipo
 * @param {string} nuevoEstado - Nuevo estado
 * @param {string} economico - Económico asignado (opcional)
 * @param {string} ticket - Ticket de garantía (opcional)
 * @param {string} observaciones - Observaciones (opcional)
 * @returns {Object} { ok: true }
 */
function actualizarEstadoInventario(token, serie, nuevoEstado, economico, ticket, observaciones) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  // Solo Admin (1) o Revisor (2) pueden cambiar estado
  if (sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para cambiar estado de inventario.' };
  }

  if (!serie) {
    return { ok: false, error: 'La serie es requerida.' };
  }

  if (!nuevoEstado) {
    return { ok: false, error: 'El nuevo estado es requerido.' };
  }

  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }

    var datos = sheet.getDataRange().getValues();
    var serieBusqueda = serie.toString().toUpperCase().trim();
    var encontrado = false;
    var filaReal = -1;

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue;

      var serieActual = fila[0].toString().toUpperCase().trim();
      if (serieActual === serieBusqueda) {
        filaReal = i + 1;
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el equipo con serie: ' + serie };
    }

    var fechaActual = new Date();
    var observacionesActual = (datos[filaReal - 1][10] || '');

    // ✅ CONSTRUIR NUEVAS OBSERVACIONES
    var nuevaObs = observacionesActual;
    if (observaciones) {
      nuevaObs = observacionesActual + (observacionesActual ? ' | ' : '') + observaciones;
    }

    // Agregar log de cambio
    var logCambio = '[' + Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '] ' +
      sesion.nombre + ' cambió estado a "' + nuevoEstado + '"';
    if (economico) logCambio += ' (Económico: ' + economico + ')';
    if (ticket) logCambio += ' (Ticket: ' + ticket + ')';
    nuevaObs = (nuevaObs ? nuevaObs + ' | ' : '') + logCambio;

    // ✅ ACTUALIZAR COLUMNAS
    sheet.getRange(filaReal, 5).setValue(nuevoEstado); // Estado

    // ✅ Manejar económico
    if (nuevoEstado === 'Instalado' && economico) {
      sheet.getRange(filaReal, 6).setValue(economico);
      sheet.getRange(filaReal, 7).setValue(fechaActual); // Fecha instalación
    } else if (nuevoEstado === 'Disponible') {
      sheet.getRange(filaReal, 6).setValue(''); // Liberar económico
      sheet.getRange(filaReal, 7).setValue(''); // Limpiar fecha instalación
    }

    // ✅ Manejar ticket de garantía
    if (nuevoEstado === 'Garantía' && ticket) {
      sheet.getRange(filaReal, 7).setValue(fechaActual); // Fecha garantía
      sheet.getRange(filaReal, 8).setValue(ticket); // Ticket garantía
    } else if (nuevoEstado === 'Disponible' || nuevoEstado === 'Instalado' || nuevoEstado === 'Baja') {
      // Si sale de garantía, limpiar ticket (opcional, depende de tu lógica)
      // sheet.getRange(filaReal, 8).setValue('');
    }

    // ✅ Actualizar observaciones
    sheet.getRange(filaReal, 10).setValue(nuevaObs);

    // ✅ Última actualización
    sheet.getRange(filaReal, 9).setValue(fechaActual);

    console.log('✅ Estado actualizado: ' + serie + ' → ' + nuevoEstado + ' por ' + sesion.nombre);

    // ✅ Si el estado es Instalado, registrar en la Bitácora
    if (nuevoEstado === 'Instalado' && economico) {
      // Buscar si hay un registro en Bitácora para este económico
      var bitacoraSheet = SHEETS.BITACORA();
      if (bitacoraSheet) {
        var bitacoraData = bitacoraSheet.getDataRange().getValues();
        var folioEncontrado = null;
        for (var b = 1; b < bitacoraData.length; b++) {
          if (bitacoraData[b][2] === economico && bitacoraData[b][12] === 'Borrador') {
            folioEncontrado = bitacoraData[b][0];
            break;
          }
        }
        // Si hay un borrador, actualizar la serie GPS
        if (folioEncontrado) {
          var idxFolio = bitacoraData[0].indexOf('FOLIO');
          for (var b = 1; b < bitacoraData.length; b++) {
            if (bitacoraData[b][idxFolio] === folioEncontrado) {
              var idxSerie = bitacoraData[0].indexOf('SERIE_GPS');
              if (idxSerie !== -1) {
                bitacoraSheet.getRange(b + 1, idxSerie + 1).setValue(serie);
                console.log('✅ Actualizada serie en Bitácora para folio:', folioEncontrado);
              }
              break;
            }
          }
        }
      }
    }

    return { ok: true };

  } catch (err) {
    console.error('Error en actualizarEstadoInventario:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// EXPORTAR INVENTARIO A EXCEL
// ============================================================

/**
 * Exporta el inventario GPS a un archivo Excel y devuelve la URL de descarga
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: boolean, url: string, error: string }
 */
function exportarInventarioExcel(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Revisor (rol 2) y Admin (rol 1) pueden exportar
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para exportar inventario.' };
  }

  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }

    // 1. OBTENER DATOS DE LA HOJA
    var datos = sheet.getDataRange().getValues();
    if (!datos || datos.length < 2) {
      return { ok: false, error: 'No hay datos en el inventario.' };
    }

    // 2. CREAR ARCHIVO HTML (Google Sheets no soporta Excel nativo, usamos HTML como tabla)
    var html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" 
                    xmlns:x="urn:schemas-microsoft-com:office:excel" 
                    xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <!--[if gte mso 9]>
      <xml>
        <x:ExcelWorkbook>
          <x:ExcelWorksheets>
            <x:ExcelWorksheet>
              <x:Name>Inventario GPS</x:Name>
              <x:WorksheetOptions>
                <x:DisplayGridlines/>
              </x:WorksheetOptions>
            </x:ExcelWorksheet>
          </x:ExcelWorksheets>
        </x:ExcelWorkbook>
      </xml>
      <![endif]-->
      <style>
        table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; }
        th { background: #1a56db; color: #ffffff; font-weight: bold; padding: 6px 10px; border: 1px solid #000; }
        td { padding: 4px 10px; border: 1px solid #ccc; }
        .header { font-size: 14pt; font-weight: bold; margin-bottom: 10px; }
        .fecha { font-size: 9pt; color: #666; margin-bottom: 15px; }
        .estado-disponible { background: #d4edda; }
        .estado-instalado { background: #cce5ff; }
        .estado-garantia { background: #fff3cd; }
        .estado-baja { background: #f8d7da; }
      </style>
    </head>
    <body>
      <div class="header">📦 INVENTARIO GPS - SAMSARA</div>
      <div class="fecha">Fecha de exportación: ${new Date().toLocaleString('es-MX')}</div>
      <table>`;

    // 3. ENCABEZADOS (usando los encabezados reales de la hoja)
    var headers = datos[0];
    html += '<thead><tr>';
    for (var h = 0; h < headers.length; h++) {
      var headerText = headers[h] || 'Columna ' + (h + 1);
      html += '<th>' + headerText + '</th>';
    }
    html += '</tr></thead><tbody>';

    // 4. DATOS (con colores según estado)
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue; // Saltar filas vacías

      var estado = (fila[4] || '').toString().trim();
      var claseEstado = '';
      if (estado === 'Disponible') claseEstado = 'estado-disponible';
      else if (estado === 'Instalado') claseEstado = 'estado-instalado';
      else if (estado === 'Garantía') claseEstado = 'estado-garantia';
      else if (estado === 'Baja') claseEstado = 'estado-baja';

      html += '<tr class="' + claseEstado + '">';
      for (var j = 0; j < fila.length; j++) {
        var valor = fila[j];
        if (valor instanceof Date) {
          valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        }
        html += '<td>' + (valor || '') + '</td>';
      }
      html += '</tr>';
    }

    html += '</tbody></table>';
    html += `<div style="margin-top: 15px; font-size: 8pt; color: #999;">
      Documento generado automáticamente por Fleet Manager - ${new Date().toLocaleString('es-MX')}
    </div>`;
    html += '</body></html>';

    // 5. CREAR ARCHIVO EN DRIVE
    var nombreArchivo = 'Inventario_GPS_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.xls';

    var blob = Utilities.newBlob(html, 'application/vnd.ms-excel', nombreArchivo);

    // Guardar en la carpeta raíz de Drive del usuario
    var file = DriveApp.createFile(blob);

    // Compartir con cualquier usuario que tenga el enlace (solo vista)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Obtener URL
    var url = file.getUrl();

    console.log('✅ Excel exportado:', nombreArchivo, 'URL:', url);

    return { ok: true, url: url };

  } catch (err) {
    console.error('❌ Error al exportar inventario:', err);
    return { ok: false, error: 'Error al exportar: ' + err.message };
  }
}
function agregarEquipoInventario(token, equipo) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  if (sesionResp.sesion.rol > 1) return { ok: false, error: 'Solo Administradores pueden agregar equipos.' };

  // TODO: Agregar el equipo a la hoja de inventario
  return { ok: true };
}
// ============================================================
// REPORTE DE UNIDADES CON SERIES (FACTURACIÓN)
// ============================================================

/**
 * Genera el reporte de unidades con series para facturación
 * Incluye el campo "Tipo" (CAJA SECA, DOLLY, MÓVIL, PRUEBA, REMOLQUE, UNIDAD, UTILITARIA)
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: boolean, url: string, error: string }
 */
function exportarReporteUnidades(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para exportar reporte.' };
  }

  try {
    // ✅ LEER VEHÍCULOS DESDE 📋_Catalogo_Vehiculos
    var catalogoSheet = SS.getSheetByName('📋_Catalogo_Vehiculos');
    if (!catalogoSheet) {
      return { ok: false, error: 'No se encontró la hoja 📋_Catalogo_Vehiculos.' };
    }

    var sheetInventario = SHEETS.INVENTARIO();
    if (!sheetInventario) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }

    // 1. OBTENER VEHÍCULOS ACTIVOS DEL CATÁLOGO
    var catalogoData = catalogoSheet.getDataRange().getValues();
    var catalogoVehiculos = {};
    
    for (var i = 1; i < catalogoData.length; i++) {
      var fila = catalogoData[i];
      if (!fila[0]) continue;
      
      var estado = (fila[8] || '').toString().trim(); // Columna I = ESTADO
      if (estado === 'Activo') {
        catalogoVehiculos[fila[0].toString().trim()] = {
          nombre: fila[0] || '',
          placas: fila[1] || '',
          tipoVehiculo: fila[2] || '',
          tipoUnidad: fila[10] || ''  // Columna K = TIPO_UNIDAD
        };
      }
    }
    console.log('📌 Vehículos activos en catálogo:', Object.keys(catalogoVehiculos).length);

    // 2. OBTENER EQUIPOS INSTALADOS O EN GARANTÍA DEL INVENTARIO
    var inventarioData = sheetInventario.getDataRange().getValues();
    var equiposPorEconomico = {};

    for (var j = 1; j < inventarioData.length; j++) {
      var fila = inventarioData[j];
      if (!fila[0]) continue;
      
      var serie = fila[0] || '';
      var tipo = fila[1] || '';
      var estado = (fila[4] || '').toString().trim();
      var economico = (fila[5] || '').toString().trim();
      
      // ✅ CORREGIDO: Incluir equipos en Garantía también (tienen económico asignado)
      // Excluir solo Baja y Disponible (sin económico)
      if (economico && estado !== 'Baja' && estado !== 'Disponible') {
        if (!equiposPorEconomico[economico]) {
          equiposPorEconomico[economico] = {
            gateway: null,
            gatewaySerie: null,
            camara: null,
            camaraSerie: null,
            accesorios: []
          };
        }
        
        var tipoUpper = tipo.toString().toUpperCase();
        if (tipoUpper.indexOf('VG') !== -1 || tipoUpper.indexOf('GATEWAY') !== -1) {
          equiposPorEconomico[economico].gateway = tipo;
          equiposPorEconomico[economico].gatewaySerie = serie;
        } else if (tipoUpper.indexOf('CM') !== -1 || tipoUpper.indexOf('CAMARA') !== -1) {
          equiposPorEconomico[economico].camara = tipo;
          equiposPorEconomico[economico].camaraSerie = serie;
        }
      }
    }

    // 3. OBTENER ACCESORIOS
    var sheetAccesorios = SHEETS.ACCESORIOS();
    if (sheetAccesorios) {
      var accesoriosData = sheetAccesorios.getDataRange().getValues();
      for (var k = 1; k < accesoriosData.length; k++) {
        var fila = accesoriosData[k];
        if (!fila[0]) continue;
        
        var tipoAccesorio = fila[1] || fila[0] || '';
        var economicoAsignado = (fila[5] || '').toString().trim();
        
        if (economicoAsignado) {
          var listaEconomicos = economicoAsignado.split(',').map(function(e) { return e.trim(); });
          listaEconomicos.forEach(function(eco) {
            if (eco && equiposPorEconomico[eco]) {
              equiposPorEconomico[eco].accesorios.push(tipoAccesorio);
            }
          });
        }
      }
    }

    // 4. GENERAR HTML PARA EXCEL
    var html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" 
                    xmlns:x="urn:schemas-microsoft-com:office:excel" 
                    xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <style>
        table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%; }
        th { background: #1a56db; color: #ffffff; font-weight: bold; padding: 8px 12px; border: 1px solid #000; text-align: left; }
        td { padding: 6px 12px; border: 1px solid #ccc; }
        .header { font-size: 14pt; font-weight: bold; margin-bottom: 10px; color: #1a56db; }
        .fecha { font-size: 9pt; color: #666; margin-bottom: 15px; }
        .footer { margin-top: 15px; font-size: 8pt; color: #999; }
        .sin-equipo { color: #999; font-style: italic; }
      </style>
    </head>
    <body>
      <div class="header">📋 REPORTE DE UNIDADES CON SERIES</div>
      <div class="fecha">Fecha de exportación: ${new Date().toLocaleString('es-MX')}</div>
      <table>
        <thead>
          <tr>
            <th>No.</th>
            <th>Nombre</th>
            <th>Dispositivo telemático</th>
            <th>Gateway Serial</th>
            <th>Etiquetas</th>
            <th>Cámara</th>
            <th>Serie de la cámara</th>
            <th>Tipo</th>
          </tr>
        </thead>
        <tbody>`;

    // 5. GENERAR FILAS (SOLO VEHÍCULOS DEL CATÁLOGO CON EQUIPOS)
    var contador = 0;
    var economicoKeys = Object.keys(equiposPorEconomico).sort();

    for (var idx = 0; idx < economicoKeys.length; idx++) {
      var eco = economicoKeys[idx];
      var equipo = equiposPorEconomico[eco];
      var infoVehiculo = catalogoVehiculos[eco] || null;
      
      contador++;
      var nombre = infoVehiculo ? infoVehiculo.nombre : eco;
      var tipoVehiculo = infoVehiculo ? (infoVehiculo.tipoUnidad || infoVehiculo.tipoVehiculo || '—') : '—';
      var etiquetas = equipo.accesorios.length > 0 ? equipo.accesorios.join(', ') : '—';
      
      // ✅ Si la cámara está en garantía pero tiene serie, mostrarla con indicador
      var camaraMostrar = equipo.camara || '—';
      var camaraSerieMostrar = equipo.camaraSerie || '—';
      
      html += `<tr>
        <td>${contador}</td>
        <td><strong>${nombre}</strong></td>
        <td>${equipo.gateway || '—'}</td>
        <td><code>${equipo.gatewaySerie || '—'}</code></td>
        <td>${etiquetas}</td>
        <td>${camaraMostrar}</td>
        <td><code>${camaraSerieMostrar}</code></td>
        <td>${tipoVehiculo}</td>
      </tr>`;
    }

    if (contador === 0) {
      html += `<tr>
        <td colspan="8" style="text-align:center;color:#999;padding:20px;">
          No hay unidades con equipos instalados.
        </td>
      </tr>`;
    }

    html += `</tbody></table>
      <div class="footer">
        Total de unidades: ${contador} | 
        Generado por Fleet Manager - ${new Date().toLocaleString('es-MX')}
      </div>
    </body></html>`;

    // 6. CREAR ARCHIVO EN DRIVE
    var nombreArchivo = 'Reporte_Unidades_Series_' + 
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.xls';
    
    var blob = Utilities.newBlob(html, 'application/vnd.ms-excel', nombreArchivo);
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = file.getUrl();
    
    console.log('✅ Reporte de unidades generado:', nombreArchivo);
    return { ok: true, url: url };

  } catch (err) {
    console.error('❌ Error al generar reporte de unidades:', err);
    return { ok: false, error: 'Error al generar reporte: ' + err.message };
  }
}
// ============================================================
// TIPOS DE UNIDAD - CATÁLOGO
// ============================================================

/**
 * Obtiene el catálogo de tipos de unidad
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, tipos: [...] }
 */
function obtenerTiposUnidad(token) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    try {
      var sheet = SS.getSheetByName('📋_Tipos_Unidad');
      if (!sheet) {
        // Si no existe la hoja, crearla con valores por defecto
        SS.insertSheet('📋_Tipos_Unidad');
        var newSheet = SS.getSheetByName('📋_Tipos_Unidad');
        newSheet.appendRow(['Tipo']);
        var tiposDefault = ['CAJA SECA', 'DOLLY', 'MÓVIL', 'PRUEBA', 'REMOLQUE', 'UNIDAD', 'UTILITARIA'];
        tiposDefault.forEach(function (t) {
          newSheet.appendRow([t]);
        });
        return { ok: true, tipos: tiposDefault };
      }

      var datos = sheet.getDataRange().getValues();
      var tipos = [];

      for (var i = 1; i < datos.length; i++) {
        var tipo = datos[i][0];
        if (tipo && tipo.toString().trim() !== '') {
          tipos.push(tipo.toString().trim().toUpperCase());
        }
      }

      return { ok: true, tipos: tipos };

    } catch (err) {
      console.error('Error al obtener tipos de unidad:', err);
      return { ok: false, error: err.message };
    }
  }
  /**
   * Agrega un nuevo tipo de unidad (solo Admin/Revisor)
   * @param {string} token - Token de sesión
   * @param {string} tipo - Nuevo tipo a agregar
   * @returns {Object} { ok: true }
   */
  function agregarTipoUnidad(token, tipo) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    if (sesionResp.sesion.rol > 2) {
      return { ok: false, error: 'Sin permisos para agregar tipos de unidad.' };
    }

    if (!tipo || tipo.trim() === '') {
      return { ok: false, error: 'El tipo de unidad es requerido.' };
    }

    try {
      var sheet = SS.getSheetByName('📋_Tipos_Unidad');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja 📋_Tipos_Unidad.' };
      }

      var tipoUpper = tipo.toString().toUpperCase().trim();

      var datos = sheet.getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        var existente = (datos[i][0] || '').toString().toUpperCase().trim();
        if (existente === tipoUpper) {
          return { ok: false, error: 'El tipo "' + tipoUpper + '" ya existe.' };
        }
      }

      sheet.appendRow([tipoUpper]);

      console.log('✅ Tipo de unidad agregado:', tipoUpper);
      return { ok: true, mensaje: 'Tipo "' + tipoUpper + '" agregado correctamente.' };

    } catch (err) {
      console.error('Error al agregar tipo de unidad:', err);
      return { ok: false, error: err.message };
    }
  }
  /**
   * Elimina un tipo de unidad (solo Admin/Revisor)
   * @param {string} token - Token de sesión
   * @param {string} tipo - Tipo a eliminar
   * @returns {Object} { ok: true }
   */
  function eliminarTipoUnidad(token, tipo) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    if (sesionResp.sesion.rol > 2) {
      return { ok: false, error: 'Sin permisos para eliminar tipos de unidad.' };
    }

    if (!tipo || tipo.trim() === '') {
      return { ok: false, error: 'El tipo de unidad es requerido.' };
    }

    try {
      var sheet = SS.getSheetByName('📋_Tipos_Unidad');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja 📋_Tipos_Unidad.' };
      }

      var tipoUpper = tipo.toString().toUpperCase().trim();
      var datos = sheet.getDataRange().getValues();
      var filaEliminar = -1;

      for (var i = 1; i < datos.length; i++) {
        var existente = (datos[i][0] || '').toString().toUpperCase().trim();
        if (existente === tipoUpper) {
          filaEliminar = i + 1;
          break;
        }
      }

      if (filaEliminar === -1) {
        return { ok: false, error: 'El tipo "' + tipoUpper + '" no existe.' };
      }

      sheet.deleteRow(filaEliminar);

      console.log('✅ Tipo de unidad eliminado:', tipoUpper);
      return { ok: true, mensaje: 'Tipo "' + tipoUpper + '" eliminado correctamente.' };

    } catch (err) {
      console.error('Error al eliminar tipo de unidad:', err);
      return { ok: false, error: err.message };
    }
  }
  /**
   * Actualiza el tipo de unidad de un vehículo en la flotilla
   * @param {string} token - Token de sesión
   * @param {string} economico - ID del vehículo
   * @param {string} tipoUnidad - Tipo de unidad seleccionado
   * @returns {Object} { ok: true }
   */
  function actualizarTipoUnidad(token, economico, tipoUnidad) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    if (sesionResp.sesion.rol > 2) {
      return { ok: false, error: 'Sin permisos para modificar tipo de unidad.' };
    }

    if (!economico) {
      return { ok: false, error: 'El económico es requerido.' };
    }

    try {
      // ✅ GUARDAR EN 📋_Catalogo_Vehiculos
      var catalogoSheet = SS.getSheetByName('📋_Catalogo_Vehiculos');
      if (!catalogoSheet) {
        return { ok: false, error: 'No se encontró la hoja 📋_Catalogo_Vehiculos.' };
      }

      var datos = catalogoSheet.getDataRange().getValues();
      var encontrado = false;
      var filaReal = -1;
      var economicoBusqueda = economico.toString().toUpperCase().trim();

      for (var i = 1; i < datos.length; i++) {
        var ecoActual = (datos[i][0] || '').toString().toUpperCase().trim();
        if (ecoActual === economicoBusqueda) {
          filaReal = i + 1;
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        // Si no existe el vehículo en el catálogo, lo creamos automáticamente
        console.warn('⚠️ Vehículo no encontrado en catálogo, creándolo...');
        var nuevaFila = [
          economico,      // ECONOMICO
          '',             // PLACAS
          '',             // TIPO_VEHICULO
          '',             // MARCA
          '',             // MODELO
          '',             // AÑO
          '',             // SERIE_VEHICULO
          '',             // GPS_ACTUAL
          'Activo',       // ESTADO
          '',             // ULTIMO_SERVICIO
          tipoUnidad      // TIPO_UNIDAD
        ];
        catalogoSheet.appendRow(nuevaFila);
        console.log('✅ Vehículo creado en catálogo:', economico);
        return { ok: true };
      }

      // ✅ Actualizar TIPO_UNIDAD en la columna 11 (índice 10)
      catalogoSheet.getRange(filaReal, 11).setValue(tipoUnidad || '');

      console.log('✅ Tipo de unidad actualizado en catálogo:', economico, '→', tipoUnidad || 'VACÍO');
      return { ok: true };

    } catch (err) {
      console.error('Error en actualizarTipoUnidad:', err);
      return { ok: false, error: err.message };
    }
  }
  /**
   * Obtiene el tipo de unidad de un vehículo específico
   * @param {string} token - Token de sesión
   * @param {string} economico - ID del vehículo
   * @returns {Object} { ok: true, tipoUnidad: string }
   */
  function obtenerTipoUnidadPorEconomico(token, economico) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    if (!economico) {
      return { ok: false, error: 'El económico es requerido.' };
    }

    try {
      // ✅ LEER DESDE 📋_Catalogo_Vehiculos
      var catalogoSheet = SS.getSheetByName('📋_Catalogo_Vehiculos');
      if (!catalogoSheet) {
        return { ok: false, error: 'No se encontró la hoja 📋_Catalogo_Vehiculos.' };
      }

      var datos = catalogoSheet.getDataRange().getValues();
      var economicoBusqueda = economico.toString().toUpperCase().trim();

      for (var i = 1; i < datos.length; i++) {
        var ecoActual = (datos[i][0] || '').toString().toUpperCase().trim();
        if (ecoActual === economicoBusqueda) {
          var tipoUnidad = (datos[i][10] || '').toString().trim(); // Columna K = TIPO_UNIDAD
          return { ok: true, tipoUnidad: tipoUnidad, existe: true };
        }
      }

      return { ok: true, tipoUnidad: '', existe: false };

    } catch (err) {
      console.error('Error en obtenerTipoUnidadPorEconomico:', err);
      return { ok: false, error: err.message };
    }
  }
  /**
   * Actualiza los datos de un equipo en el inventario
   * @param {string} token - Token de sesión
   * @param {string} serie - Serie del equipo
   * @param {Object} datos - Datos a actualizar { tipo, modelo, imei, estado, economico, observaciones }
   * @returns {Object} { ok: true }
   */
  function actualizarEquipoInventario(token, serie, datos) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    // Solo Admin (1) y Revisor (2) pueden editar
    if (sesionResp.sesion.rol > 2) {
      return { ok: false, error: 'Sin permisos para editar equipos.' };
    }

    if (!serie) {
      return { ok: false, error: 'La serie es requerida.' };
    }

    try {
      var sheet = SHEETS.INVENTARIO();
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja de inventario.' };
      }

      var datosSheet = sheet.getDataRange().getValues();
      var serieBusqueda = serie.toString().toUpperCase().trim();
      var encontrado = false;
      var filaReal = -1;

      for (var i = 1; i < datosSheet.length; i++) {
        var serieActual = (datosSheet[i][0] || '').toString().toUpperCase().trim();
        if (serieActual === serieBusqueda) {
          filaReal = i + 1;
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        return { ok: false, error: 'No se encontró el equipo con serie: ' + serie };
      }

      // ✅ ACTUALIZAR CAMPOS
      // Columna 2 (B): TIPO_EQUIPO
      if (datos.tipo !== undefined) {
        sheet.getRange(filaReal, 2).setValue(datos.tipo);
      }

      // Columna 3 (C): MODELO
      if (datos.modelo !== undefined) {
        sheet.getRange(filaReal, 3).setValue(datos.modelo);
      }

      // Columna 4 (D): IMEI
      if (datos.imei !== undefined) {
        sheet.getRange(filaReal, 4).setValue(datos.imei);
      }

      // Columna 5 (E): ESTADO
      if (datos.estado !== undefined) {
        sheet.getRange(filaReal, 5).setValue(datos.estado);
      }

      // Columna 6 (F): ECONOMICO_ASIGNADO
      if (datos.economico !== undefined) {
        sheet.getRange(filaReal, 6).setValue(datos.economico);

        // Si el estado cambió a Instalado y hay económico, actualizar fecha
        if (datos.estado === 'Instalado' && datos.economico) {
          sheet.getRange(filaReal, 7).setValue(new Date()); // FECHA_INSTALACION
        } else if (datos.estado === 'Disponible') {
          // Si se libera, limpiar económico y fecha
          sheet.getRange(filaReal, 6).setValue('');
          sheet.getRange(filaReal, 7).setValue('');
        }
      }

      // Columna 11 (K): OBSERVACIONES
      if (datos.observaciones !== undefined) {
        var obsActual = (datosSheet[filaReal - 1][10] || '');
        var nuevaObs = obsActual + (obsActual ? ' | ' : '') + '[EDIT] ' + datos.observaciones;
        sheet.getRange(filaReal, 11).setValue(nuevaObs);
      }

      // Columna 10 (J): ULTIMA_ACTUALIZACION
      sheet.getRange(filaReal, 10).setValue(new Date());

      console.log('✅ Equipo actualizado:', serie);
      return { ok: true };

    } catch (err) {
      console.error('Error en actualizarEquipoInventario:', err);
      return { ok: false, error: err.message };
    }
  }
  // ============================================================
  // TICKETS - BACKEND
  // ============================================================

  /**
   * Genera un nuevo ID de ticket secuencial
   */
  function generarIdTicket() {
    var params = _leerParams();
    var prefijo = params['TICKET_PREFIJO'] || 'TKT';
    var ultimo = parseInt(params['TICKET_ULTIMO'] || '0', 10);
    var nuevo = ultimo + 1;
    var ticketId = prefijo + '-' + String(nuevo).padStart(4, '0');
    _escribirParam('TICKET_ULTIMO', String(nuevo));
    return ticketId;
  }

  /**
   * Obtiene todos los tickets
   * @param {string} token - Token de sesión
   * @param {Object} filtros - Filtros opcionales { estado, unidad, tecnico }
   * @returns {Object} { ok: true, tickets: [...] }
   */
  // ============================================================
  // TICKETS - BACKEND
  // ============================================================
  function pruebaTickets() {
    return { ok: true, mensaje: 'Función de prueba funcionando' };
  }

  /**
   * Obtiene todos los tickets
   * @param {string} token - Token de sesión
   * @param {Object} filtros - Filtros opcionales { estado, unidad, tecnico }
   * @returns {Object} { ok: true, tickets: [...] }
   */
  function obtenerTickets(token, filtros) {
    console.log('🎫 obtenerTickets - INICIO SIMPLIFICADO');

    try {
      // 1. Obtener la hoja
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('🎫_Tickets');

      if (!sheet) {
        console.warn('⚠️ No se encontró la hoja 🎫_Tickets');
        return { ok: true, tickets: [] };
      }

      // 2. Obtener datos
      var datos = sheet.getDataRange().getValues();
      var tickets = [];

      // 3. Recorrer filas (desde la 2, saltando encabezado)
      for (var i = 1; i < datos.length; i++) {
        var fila = datos[i];
        if (!fila || !fila[0]) continue; // Saltar filas vacías

        var ticket = {
          id: String(fila[0] || ''),
          fecha: String(fila[1] || ''),
          unidad: String(fila[2] || ''),
          descripcion: String(fila[3] || ''),
          creadoPor: String(fila[4] || ''),
          creadoPorNombre: String(fila[5] || ''),
          estado: String(fila[6] || 'Pendiente'),
          tecnicoAsignado: String(fila[7] || ''),
          tecnicoNombre: String(fila[8] || ''),
          fechaCierre: String(fila[9] || ''),
          comentarios: String(fila[10] || ''),
          ultimaActualizacion: String(fila[11] || '')
        };

        tickets.push(ticket);
      }

      console.log('✅ Tickets encontrados:', tickets.length);
      return { ok: true, tickets: tickets };

    } catch (err) {
      console.error('❌ Error en obtenerTickets:', err);
      return { ok: false, error: String(err.message) };
    }
  }
  function diagnosticarTickets() {
    try {
      console.log('🔍 DIAGNÓSTICO DE TICKETS');

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('🎫_Tickets');
      console.log('📌 Hoja 🎫_Tickets existe?', sheet !== null);

      if (!sheet) {
        return { ok: false, error: 'La hoja 🎫_Tickets no existe' };
      }

      var datos = sheet.getDataRange().getValues();
      console.log('📌 Filas:', datos.length);
      console.log('📌 Encabezado:', datos[0]);

      return {
        ok: true,
        existe: true,
        filas: datos.length,
        encabezado: datos[0] || [],
        primeraFila: datos[1] || []
      };

    } catch (err) {
      console.error('❌ Error en diagnóstico:', err);
      return { ok: false, error: err.message };
    }
  }

  /**
   * Genera un nuevo ID de ticket secuencial
   */
  function generarIdTicket() {
    var params = _leerParams();
    var prefijo = params['TICKET_PREFIJO'] || 'TKT';
    var ultimo = parseInt(params['TICKET_ULTIMO'] || '0', 10);
    var nuevo = ultimo + 1;
    var ticketId = prefijo + '-' + String(nuevo).padStart(4, '0');
    _escribirParam('TICKET_ULTIMO', String(nuevo));
    return ticketId;
  }

  /**
   * Crea un nuevo ticket
   */
  function crearTicket(token, datos) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
    var sesion = sesionResp.sesion;

    if (!datos.unidad || !datos.descripcion) {
      return { ok: false, error: 'La unidad y la descripción son requeridas.' };
    }

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('🎫_Tickets');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja de tickets.' };
      }

      var ticketId = generarIdTicket();
      var ahora = new Date();

      var nuevaFila = [
        ticketId,
        ahora,
        datos.unidad,
        datos.descripcion,
        sesion.usuarioId,
        sesion.nombre,
        'Pendiente',
        '',
        '',
        '',
        '',
        ahora
      ];

      sheet.appendRow(nuevaFila);

      console.log('✅ Ticket creado:', ticketId);
      return { ok: true, id: ticketId };

    } catch (err) {
      console.error('Error en crearTicket:', err);
      return { ok: false, error: err.message };
    }
  }

  /**
   * Asigna un ticket a un técnico (tomar ticket)
   */
  function tomarTicket(token, ticketId) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
    var sesion = sesionResp.sesion;

    if (!ticketId) {
      return { ok: false, error: 'El ID del ticket es requerido.' };
    }

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('🎫_Tickets');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja de tickets.' };
      }

      var datos = sheet.getDataRange().getValues();
      var encontrado = false;
      var filaReal = -1;
      var estadoActual = '';

      for (var i = 1; i < datos.length; i++) {
        if ((datos[i][0] || '').toString().trim() === ticketId) {
          filaReal = i + 1;
          estadoActual = (datos[i][6] || '').toString().trim();
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        return { ok: false, error: 'No se encontró el ticket: ' + ticketId };
      }

      if (estadoActual === 'Resuelto') {
        return { ok: false, error: 'El ticket ya está resuelto.' };
      }

      sheet.getRange(filaReal, 6).setValue('En proceso');
      sheet.getRange(filaReal, 7).setValue(sesion.usuarioId);
      sheet.getRange(filaReal, 8).setValue(sesion.nombre);
      sheet.getRange(filaReal, 11).setValue(new Date());

      console.log('✅ Ticket tomado:', ticketId, 'por', sesion.nombre);
      return { ok: true };

    } catch (err) {
      console.error('Error en tomarTicket:', err);
      return { ok: false, error: err.message };
    }
  }

  /**
   * Crea un nuevo ticket
   * @param {string} token - Token de sesión
   * @param {Object} datos - { unidad, descripcion }
   * @returns {Object} { ok: true, id: string }
   */
  function crearTicket(token, datos) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
    var sesion = sesionResp.sesion;

    if (!datos.unidad || !datos.descripcion) {
      return { ok: false, error: 'La unidad y la descripción son requeridas.' };
    }

    try {
      var sheet = SS.getSheetByName('🎫_Tickets');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja de tickets.' };
      }

      var ticketId = generarIdTicket();
      var ahora = new Date();

      var nuevaFila = [
        ticketId,
        ahora,
        datos.unidad,
        datos.descripcion,
        sesion.usuarioId,
        sesion.nombre,
        'Pendiente',
        '',
        '',
        '',
        '',
        ahora
      ];

      sheet.appendRow(nuevaFila);

      console.log('✅ Ticket creado:', ticketId);
      return { ok: true, id: ticketId };

    } catch (err) {
      console.error('Error en crearTicket:', err);
      return { ok: false, error: err.message };
    }
  }

  /**
   * Asigna un ticket a un técnico (tomar ticket)
   * @param {string} token - Token de sesión
   * @param {string} ticketId - ID del ticket
   * @returns {Object} { ok: true }
   */
  function tomarTicket(token, ticketId) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
    var sesion = sesionResp.sesion;

    if (!ticketId) {
      return { ok: false, error: 'El ID del ticket es requerido.' };
    }

    try {
      var sheet = SS.getSheetByName('🎫_Tickets');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja de tickets.' };
      }

      var datos = sheet.getDataRange().getValues();
      var encontrado = false;
      var filaReal = -1;
      var estadoActual = '';

      for (var i = 1; i < datos.length; i++) {
        if ((datos[i][0] || '').toString().trim() === ticketId) {
          filaReal = i + 1;
          estadoActual = (datos[i][6] || '').toString().trim();
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        return { ok: false, error: 'No se encontró el ticket: ' + ticketId };
      }

      if (estadoActual === 'Resuelto') {
        return { ok: false, error: 'El ticket ya está resuelto.' };
      }

      // Actualizar: Estado → En proceso, Técnico → usuario actual
      sheet.getRange(filaReal, 6).setValue('En proceso');
      sheet.getRange(filaReal, 7).setValue(sesion.usuarioId);   // TECNICO_ASIGNADO
      sheet.getRange(filaReal, 8).setValue(sesion.nombre);       // TECNICO_NOMBRE
      sheet.getRange(filaReal, 11).setValue(new Date());         // ULTIMA_ACTUALIZACION

      console.log('✅ Ticket tomado:', ticketId, 'por', sesion.nombre);
      return { ok: true };

    } catch (err) {
      console.error('Error en tomarTicket:', err);
      return { ok: false, error: err.message };
    }
  }

  /**
   * Marca un ticket como resuelto
   * @param {string} token - Token de sesión
   * @param {string} ticketId - ID del ticket
   * @param {string} comentarios - Comentarios de la solución
   * @returns {Object} { ok: true }
   */
  function resolverTicket(token, ticketId, comentarios) {
    console.log('🔧 resolverTicket - INICIO');
    console.log('📌 ticketId:', ticketId);
    console.log('📌 comentarios:', comentarios);

    try {
      // Validar sesión
      var sesionResp = validarSesion(token);
      if (!sesionResp.ok) {
        return { ok: false, error: 'Sesión inválida: ' + sesionResp.error };
      }
      var sesion = sesionResp.sesion;
      console.log('✅ Sesión válida - Usuario:', sesion.nombre);

      if (!ticketId) {
        return { ok: false, error: 'El ID del ticket es requerido.' };
      }

      // Obtener la hoja
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('🎫_Tickets');
      if (!sheet) {
        return { ok: false, error: 'No se encontró la hoja de tickets.' };
      }

      var datos = sheet.getDataRange().getValues();
      var encontrado = false;
      var filaReal = -1;
      var estadoActual = '';

      for (var i = 1; i < datos.length; i++) {
        var idActual = (datos[i][0] || '').toString().trim();
        if (idActual === ticketId) {
          filaReal = i + 1;
          estadoActual = (datos[i][6] || '').toString().trim();
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        return { ok: false, error: 'No se encontró el ticket: ' + ticketId };
      }

      if (estadoActual === 'Resuelto') {
        return { ok: false, error: 'El ticket ya está resuelto.' };
      }

      // ✅ Actualizar: Estado → Resuelto
      sheet.getRange(filaReal, 7).setValue('Resuelto'); // Columna G: ESTADO

      // ✅ Actualizar fecha de cierre
      sheet.getRange(filaReal, 10).setValue(new Date()); // Columna J: FECHA_CIERRE

      // ✅ Agregar comentarios
      if (comentarios) {
        var comentariosActual = (datos[filaReal - 1][10] || '');
        var nuevosComentarios = comentariosActual + (comentariosActual ? '\n' : '') +
          '[' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '] ' +
          sesion.nombre + ': ' + comentarios;
        sheet.getRange(filaReal, 11).setValue(nuevosComentarios); // Columna K: COMENTARIOS
      }

      // ✅ Actualizar última modificación
      sheet.getRange(filaReal, 12).setValue(new Date()); // Columna L: ULTIMA_ACTUALIZACION

      console.log('✅ Ticket resuelto:', ticketId);
      return { ok: true };

    } catch (err) {
      console.error('❌ Error en resolverTicket:', err);
      return { ok: false, error: err.message };
    }
  }




// ============================================================
// FLOTILLA - FUNCIONES DEL SERVIDOR (CORREGIDAS)
// ============================================================

/**
 * Obtiene la flotilla completa con todos los vehículos y sus equipos
 * @param {string} token - Token de autenticación
 * @param {object} filtros - Filtros opcionales { estado, tipoUnidad, buscar }
 * @returns {object} { ok, vehiculos: [...] }
 */
function obtenerFlotillaCompleta(token, filtros) {
  console.log('🚚 obtenerFlotillaCompleta - INICIO');
  
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    console.error('❌ Sesión inválida:', sesionResp.error);
    return { ok: false, error: sesionResp.error };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ✅ LEER DESDE 📋_Catalogo_Vehiculos
    var sheetVehiculos = ss.getSheetByName('📋_Catalogo_Vehiculos');
    if (!sheetVehiculos) {
      console.error('❌ No se encontró la hoja 📋_Catalogo_Vehiculos');
      return { ok: false, error: 'No se encontró la hoja 📋_Catalogo_Vehiculos.' };
    }
    
    var dataVehiculos = sheetVehiculos.getDataRange().getValues();
    console.log('📌 Filas en catálogo:', dataVehiculos.length);
    
    if (dataVehiculos.length <= 1) {
      console.log('📋 No hay vehículos en el catálogo');
      return { ok: true, vehiculos: [] };
    }
    
    // ✅ MAPEAR ENCABEZADOS CORRECTAMENTE
    var headers = dataVehiculos[0];
    var idxEconomico = headers.indexOf('ECONOMICO');
    var idxPlacas = headers.indexOf('PLACAS');
    var idxTipoVehiculo = headers.indexOf('TIPO_VEHICULO');
    var idxMarca = headers.indexOf('MARCA');
    var idxModelo = headers.indexOf('MODELO');
    var idxAnio = headers.indexOf('AÑO');
    var idxEstado = headers.indexOf('ESTADO');
    var idxTipoUnidad = headers.indexOf('TIPO_UNIDAD');
    
    console.log('📌 Índices encontrados:', { idxEconomico, idxPlacas, idxTipoVehiculo, idxMarca, idxModelo, idxAnio, idxEstado, idxTipoUnidad });
    
    // ✅ OBTENER EQUIPOS DEL INVENTARIO
    var sheetEquipos = ss.getSheetByName('📦_Inventario_GPS');
    var equiposPorEconomico = {};
    
    if (sheetEquipos) {
      var dataEquipos = sheetEquipos.getDataRange().getValues();
      var headersEquipos = dataEquipos[0];
      
      var idxEqSerie = headersEquipos.indexOf('SERIE_GPS');
      var idxEqTipo = headersEquipos.indexOf('TIPO_EQUIPO');
      var idxEqEconomico = headersEquipos.indexOf('ECONOMICO_ASIGNADO');
      var idxEqEstado = headersEquipos.indexOf('ESTADO');
      
      console.log('📌 Índices de equipos:', { idxEqSerie, idxEqTipo, idxEqEconomico, idxEqEstado });
      
      for (var i = 1; i < dataEquipos.length; i++) {
        var row = dataEquipos[i];
        var economico = row[idxEqEconomico] || '';
        var tipo = row[idxEqTipo] || '';
        var serie = row[idxEqSerie] || '';
        var estado = row[idxEqEstado] || '';
        
        if (economico && (estado === 'Instalado' || estado === 'Garantía')) {
          if (!equiposPorEconomico[economico]) {
            equiposPorEconomico[economico] = { gateway: null, gatewaySerie: null, camara: null, camaraSerie: null };
          }
          
          var tipoUpper = tipo.toString().toUpperCase();
          if (tipoUpper.indexOf('VG') !== -1 || tipoUpper.indexOf('GATEWAY') !== -1) {
            equiposPorEconomico[economico].gateway = tipo;
            equiposPorEconomico[economico].gatewaySerie = serie;
          } else if (tipoUpper.indexOf('CM') !== -1 || tipoUpper.indexOf('CAMARA') !== -1) {
            equiposPorEconomico[economico].camara = tipo;
            equiposPorEconomico[economico].camaraSerie = serie;
          }
        }
      }
      console.log('📌 Equipos por económico:', Object.keys(equiposPorEconomico));
    }
    
    // ✅ OBTENER ACCESORIOS
    var sheetAccesorios = ss.getSheetByName('🔧_Accesorios_Stock');
    var accesoriosPorEconomico = {};
    
    if (sheetAccesorios) {
      var dataAccesorios = sheetAccesorios.getDataRange().getValues();
      var headersAccesorios = dataAccesorios[0];
      var idxAccTipo = headersAccesorios.indexOf('TIPO');
      var idxAccEconomico = headersAccesorios.indexOf('ECONOMICO_ASIGNADO');
      
      for (var i = 1; i < dataAccesorios.length; i++) {
        var row = dataAccesorios[i];
        var economicoAsignado = row[idxAccEconomico] || '';
        var tipoAccesorio = row[idxAccTipo] || '';
        
        if (economicoAsignado && tipoAccesorio) {
          var listaEconomicos = economicoAsignado.split(',').map(function(e) { return e.trim(); });
          listaEconomicos.forEach(function(eco) {
            if (eco) {
              if (!accesoriosPorEconomico[eco]) {
                accesoriosPorEconomico[eco] = [];
              }
              if (accesoriosPorEconomico[eco].indexOf(tipoAccesorio) === -1) {
                accesoriosPorEconomico[eco].push(tipoAccesorio);
              }
            }
          });
        }
      }
    }
    
    // ✅ APLICAR FILTROS
    var filtroEstado = (filtros && filtros.estado) ? filtros.estado.toString().trim() : '';
    var filtroTipoUnidad = (filtros && filtros.tipoUnidad) ? filtros.tipoUnidad.toString().trim() : '';
    var filtroBuscar = (filtros && filtros.buscar) ? filtros.buscar.toString().toUpperCase().trim() : '';
    
    // ✅ CONSTRUIR RESULTADO
    var vehiculos = [];
    
    for (var i = 1; i < dataVehiculos.length; i++) {
      var row = dataVehiculos[i];
      var economico = (row[idxEconomico] || '').toString().trim();
      
      if (!economico) continue;
      
      var estado = (row[idxEstado] || '').toString().trim();
      var tipoUnidad = (row[idxTipoUnidad] || '').toString().trim();
      var placas = (row[idxPlacas] || '').toString().trim();
      var tipoVehiculo = (row[idxTipoVehiculo] || '').toString().trim();
      var marca = (row[idxMarca] || '').toString().trim();
      var modelo = (row[idxModelo] || '').toString().trim();
      var anio = (row[idxAnio] || '').toString().trim();
      
      // Aplicar filtros
      if (filtroEstado && estado !== filtroEstado) continue;
      if (filtroTipoUnidad && tipoUnidad !== filtroTipoUnidad) continue;
      if (filtroBuscar) {
        var ecoUpper = economico.toUpperCase();
        var placasUpper = placas.toUpperCase();
        if (ecoUpper.indexOf(filtroBuscar) === -1 && placasUpper.indexOf(filtroBuscar) === -1) {
          continue;
        }
      }
      
      var equipos = equiposPorEconomico[economico] || { gateway: null, gatewaySerie: null, camara: null, camaraSerie: null };
      var accesorios = accesoriosPorEconomico[economico] || [];
      
      vehiculos.push({
        economico: economico,
        placas: placas || '—',
        tipoVehiculo: tipoVehiculo || '—',
        marca: marca || '',
        modelo: modelo || '',
        anio: anio || '',
        estado: estado || 'Inactivo',
        tipoUnidad: tipoUnidad || '—',
        gateway: equipos.gateway || null,
        gatewaySerie: equipos.gatewaySerie || null,
        camara: equipos.camara || null,
        camaraSerie: equipos.camaraSerie || null,
        accesorios: accesorios || [],
        tieneEquipos: !!(equipos.gateway || equipos.camara || accesorios.length > 0)
      });
    }
    
    console.log('✅ Vehículos encontrados:', vehiculos.length);
    return { ok: true, vehiculos: vehiculos };
    
  } catch (error) {
    console.error('❌ Error en obtenerFlotillaCompleta:', error);
    return { ok: false, error: error.toString() };
  }
}
/**
 * Agrega un nuevo vehículo al catálogo
 * @param {string} token - Token de autenticación
 * @param {object} datos - Datos del vehículo
 * @returns {object} { ok, mensaje }
 */
function agregarVehiculoCatalogo(token, datos) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }
  var sesion = sesionResp.sesion;

  if (sesion.rol > 2) {
    return { ok: false, error: 'No tienes permisos para agregar vehículos.' };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ✅ CORREGIDO: Usar 📋_Catalogo_Vehiculos
    var sheet = ss.getSheetByName('📋_Catalogo_Vehiculos');
    
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja 📋_Catalogo_Vehiculos.' };
    }
    
    // Verificar si ya existe el económico
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim() === datos.economico) {
        return { ok: false, error: 'El económico "' + datos.economico + '" ya existe.' };
      }
    }
    
    // Agregar nueva fila
    var nuevaFila = [
      datos.economico,           // ECONOMICO
      datos.placas,              // PLACAS
      datos.tipoVehiculo,        // TIPO_VEHICULO
      datos.marca,               // MARCA
      datos.modelo,              // MODELO
      datos.anio,                // ANIO
      datos.serieVehiculo || '', // SERIE_VEHICULO
      datos.gpsActual || '',     // GPS_ACTUAL
      datos.estado,              // ESTADO
      new Date(),                // FECHA_REGISTRO
      datos.tipoUnidad || ''     // TIPO_UNIDAD
    ];
    
    sheet.appendRow(nuevaFila);
    
    return { ok: true, mensaje: 'Vehículo ' + datos.economico + ' agregado correctamente.' };
    
  } catch (error) {
    console.error('Error en agregarVehiculoCatalogo:', error);
    return { ok: false, error: error.toString() };
  }
}
/**
 * Obtiene los tipos de vehículo disponibles
 * @param {string} token - Token de autenticación
 * @returns {object} { ok, tipos: [...] }
 */
function obtenerTiposVehiculo(token) {
  // ✅ USAR validarSesion
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Tipos_Vehiculo');
    
    if (!sheet) {
      // Si no existe la hoja, devolver valores por defecto
      return { ok: true, tipos: ['Camión', 'Pickup', 'Van', 'SUV', 'Sedán', 'Otro'] };
    }
    
    var data = sheet.getDataRange().getValues();
    var tipos = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        tipos.push(data[i][0].toString().trim());
      }
    }
    
    return { ok: true, tipos: tipos };
    
  } catch (error) {
    console.error('Error en obtenerTiposVehiculo:', error);
    return { ok: false, error: error.toString() };
  }
}
/**
 * Obtiene los estados de vehículo disponibles
 * @param {string} token - Token de autenticación
 * @returns {object} { ok, estados: [...] }
 */
function obtenerEstadosVehiculo(token) {
  // ✅ USAR validarSesion
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }
  
  // Estados predefinidos (no necesita hoja separada)
  return { ok: true, estados: ['Activo', 'Inactivo', 'En mantenimiento'] };
}
/**
 * Función de depuración para verificar equipos de un vehículo
 */
function diagnosticarEquiposVehiculo(token, economico) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetEquipos = ss.getSheetByName('📦_Inventario_GPS');
    
    if (!sheetEquipos) {
      return { ok: false, error: 'No se encontró la hoja de inventario.' };
    }
    
    var datos = sheetEquipos.getDataRange().getValues();
    var headers = datos[0];
    
    var idxSerie = headers.indexOf('SERIE');
    var idxTipo = headers.indexOf('TIPO');
    var idxEstado = headers.indexOf('ESTADO');
    var idxEconomico = headers.indexOf('ECONOMICO');
    
    var resultados = [];
    var economicoBusqueda = economico.toString().trim();
    
    for (var i = 1; i < datos.length; i++) {
      var row = datos[i];
      var ecoActual = (row[idxEconomico] || '').toString().trim();
      var estado = (row[idxEstado] || '').toString().trim();
      var serie = row[idxSerie] || '';
      var tipo = row[idxTipo] || '';
      
      resultados.push({
        serie: serie,
        tipo: tipo,
        estado: estado,
        economico: ecoActual,
        coincide: ecoActual === economicoBusqueda
      });
    }
    
    return {
      ok: true,
      economicoBuscado: economicoBusqueda,
      totalEquipos: resultados.length,
      equipos: resultados,
      equiposCoincidentes: resultados.filter(function(r) { return r.coincide && r.estado === 'Instalado'; })
    };
    
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
/**
 * Actualiza los datos de un vehículo en el catálogo
 * @param {string} token - Token de sesión
 * @param {string} economico - ID del vehículo
 * @param {Object} datos - Datos a actualizar
 * @returns {Object} { ok: true }
 */
function actualizarVehiculoCatalogo(token, economico, datos) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para editar vehículos.' };
  }

  if (!economico) {
    return { ok: false, error: 'El económico es requerido.' };
  }

  try {
    var sheet = SS.getSheetByName('📋_Catalogo_Vehiculos');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja 📋_Catalogo_Vehiculos.' };
    }

    var datosSheet = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;
    var economicoBusqueda = economico.toString().toUpperCase().trim();

    for (var i = 1; i < datosSheet.length; i++) {
      var ecoActual = (datosSheet[i][0] || '').toString().toUpperCase().trim();
      if (ecoActual === economicoBusqueda) {
        filaReal = i + 1;
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el vehículo con económico: ' + economico };
    }

    // Actualizar columnas
    if (datos.placas !== undefined) sheet.getRange(filaReal, 2).setValue(datos.placas);
    if (datos.tipoVehiculo !== undefined) sheet.getRange(filaReal, 3).setValue(datos.tipoVehiculo);
    if (datos.marca !== undefined) sheet.getRange(filaReal, 4).setValue(datos.marca);
    if (datos.modelo !== undefined) sheet.getRange(filaReal, 5).setValue(datos.modelo);
    if (datos.anio !== undefined) sheet.getRange(filaReal, 6).setValue(datos.anio);
    if (datos.estado !== undefined) sheet.getRange(filaReal, 9).setValue(datos.estado);
    if (datos.tipoUnidad !== undefined) sheet.getRange(filaReal, 11).setValue(datos.tipoUnidad);

    console.log('✅ Vehículo actualizado:', economico);
    return { ok: true };

  } catch (err) {
    console.error('Error en actualizarVehiculoCatalogo:', err);
    return { ok: false, error: err.message };
  }
}