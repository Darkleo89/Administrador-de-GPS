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
  NOTIFICACIONES: () => SS.getSheetByName('📩_Notificaciones'),
  TICKETS: () => SS.getSheetByName('🎫_Tickets'),
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

        // ✅ REGISTRAR LOGIN EXITOSO EN AUDITORÍA
        _registrarAuditoria(
          usuarioId.toString(),
          nombre.toString(),
          'LOGIN_EXITOSO',
          'AUTENTICACION',
          'Inicio de sesión exitoso para ' + user,
          '',
          '',
          ''
        );

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

    // ✅ REGISTRAR LOGIN FALLIDO EN AUDITORÍA
    _registrarAuditoria(
      '',
      username,
      'LOGIN_FALLIDO',
      'AUTENTICACION',
      'Intento de inicio de sesión fallido para usuario: ' + username,
      '',
      '',
      ''
    );

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
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    texto,
    Utilities.Charset.UTF_8
  );
  var resultado = '';
  for (var i = 0; i < bytes.length; i++) {
    resultado += ('0' + (bytes[i] & 0xff).toString(16)).slice(-2);
  }
  return resultado;
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
/**
 * Cierra la sesión del usuario
 * @param {string} token - Token de sesión a invalidar
 * @returns {Object} Resultado de la operación
 */
function cerrarSesion(token) {
  try {
    if (!token) {
      return { ok: false, error: 'Token no proporcionado' };
    }
    
    // Opcional: Invalidar token en la base de datos
    // _invalidarToken(token);
    
    console.log('🔐 Sesión cerrada:', token.substring(0, 10) + '...');
    return { ok: true, mensaje: 'Sesión cerrada exitosamente' };
    
  } catch (err) {
    console.error('❌ Error al cerrar sesión:', err);
    return { ok: false, error: err.message };
  }
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
      // ✅ AGREGADO: Actualizar fecha de modificación
      sheet.getRange(i + 1, 4).setValue(new Date());
      return;
    }
  }

  // ✅ AGREGADO: Si no existe la clave, crearla
  const nuevaFila = sheet.getLastRow() + 1;
  sheet.getRange(nuevaFila, 1).setValue(clave);
  sheet.getRange(nuevaFila, 2).setValue(valor);
  sheet.getRange(nuevaFila, 4).setValue(new Date());
  console.log('✅ Nuevo parámetro creado:', clave, '=', valor);
}


// ════════════════════════════════════════════════════════════
// 4. RECEPCIÓN DE REPORTE (Orden de Servicio)
// ════════════════════════════════════════════════════════════


/**
 * Punto de entrada principal mejorado para recibir, procesar y guardar reportes técnicos.
 * Soporta la creación de registros nuevos y la actualización de borradores existentes.
 * 
 * @param {string} token - Token de sesión del técnico
 * @param {Object} datos - Objeto con los campos del formulario de instalación
 * @param {Array} archivos - Arreglo de fotos evidencias en Base64
 * @returns {Object} { ok: boolean, folio: string, folderUrl: string, mensaje: string }
 */
/**
 * Punto de entrada principal para recibir, procesar y guardar reportes técnicos.
 * Soporta la creación de registros nuevos y la actualización de borradores existentes.
 * 
 * @param {string} token - Token de sesión del técnico
 * @param {Object} datos - Objeto con los campos del formulario de instalación
 * @param {Array} archivos - Arreglo de fotos evidencias en Base64
 * @returns {Object} { ok: boolean, folio: string, folderUrl: string, mensaje: string }
 */
function recibirReporte(token, datos, archivos) {
  // 1. Validación de sesión
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error || 'Sesión inválida o expirada' };
  }
  
  const sesion = sesionResp.sesion;

  // Validación de permisos
  if (sesion.rol > 3) {
    return { ok: false, error: 'Sin permisos para reportar.' };
  }

  try {
    // 2. Validar hoja de Bitácora
    if (typeof SHEETS === 'undefined' || !SHEETS.BITACORA) {
      return { ok: false, error: 'Configuración del sistema incompleta.' };
    }

    const sheet = SHEETS.BITACORA();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Bitácora.' };
    }

    // 3. Leer datos de la bitácora
    const ahora = new Date();
    const datosBitacora = sheet.getDataRange().getValues();
    
    if (datosBitacora.length === 0) {
      return { ok: false, error: 'La hoja de Bitácora está vacía o no tiene encabezados.' };
    }
    
    const headers = datosBitacora[0];

    // Localizar índices de columnas
    const idxFolio = headers.indexOf('FOLIO');
    const idxEstado = headers.indexOf('ESTADO');
    const idxTecnico = headers.indexOf('TECNICO');
    const idxFecha = headers.indexOf('FECHA_REPORTE');
    const idxDatos = headers.indexOf('DATOS_JSON');
    const idxFotos = headers.indexOf('FOTOS_DRIVE_URL');
    const idxGPS = headers.indexOf('GPS');

    if (idxFolio === -1) {
      return { ok: false, error: 'La hoja de Bitácora no tiene la columna "FOLIO".' };
    }

    // 4. Detectar si es edición
    let folio = datos.folioEdicion || datos.folio || null;
    let filaDestino = -1;
    let folderUrl = '';
    let esNuevo = false;
    let registroExistente = null;

    if (folio) {
      const folioBusqueda = folio.toString().toUpperCase().trim();
      for (let i = 1; i < datosBitacora.length; i++) {
        const folioActual = (datosBitacora[i][idxFolio] || '').toString().toUpperCase().trim();
        if (folioActual === folioBusqueda) {
          filaDestino = i + 1;
          registroExistente = datosBitacora[i];
          if (idxFotos !== -1) {
            folderUrl = datosBitacora[i][idxFotos] || '';
          }
          break;
        }
      }
    }

    // 5. Procesar fotos
    if (filaDestino === -1) {
      // NUEVO REPORTE
      esNuevo = true;
      folio = generarFolio();
      
      const resultadoFotos = _subirFotosDrive(folio, archivos);
      if (typeof resultadoFotos === 'string') {
        folderUrl = resultadoFotos;
      } else if (resultadoFotos && typeof resultadoFotos === 'object') {
        folderUrl = resultadoFotos.folderUrl || resultadoFotos.url || '';
      } else {
        folderUrl = '';
      }
      
      console.log(`📁 Carpeta creada para nuevo reporte ${folio}`);
      
    } else {
      // EDICIÓN DE REPORTE EXISTENTE
      esNuevo = false;
      
      if (archivos && Array.isArray(archivos) && archivos.length > 0) {
        if (folderUrl && folderUrl !== '') {
          console.log(`📸 Anexando ${archivos.length} fotos al reporte ${folio}...`);
          
          // ✅ FUNCIÓN CORREGIDA Y VALIDADA
          const resultadoAnexo = _anexarFotosACarpetaExistente(folderUrl, archivos, {
            organizarPorFecha: true
          });
          
          if (resultadoAnexo.success) {
            if (resultadoAnexo.folderUrl) {
              folderUrl = resultadoAnexo.folderUrl;
            }
            console.log(`✅ ${resultadoAnexo.archivosSubidos} fotos anexadas correctamente`);
          } else {
            console.warn(`⚠️ Error al anexar fotos:`, resultadoAnexo.errores);
          }
        } else {
          // Crear nueva carpeta si no existe
          const resultadoFotos = _subirFotosDrive(folio, archivos);
          if (typeof resultadoFotos === 'string') {
            folderUrl = resultadoFotos;
          } else if (resultadoFotos && typeof resultadoFotos === 'object') {
            folderUrl = resultadoFotos.folderUrl || resultadoFotos.url || '';
          }
        }
      }
    }

    // 6. Preparar y guardar datos
    const datosCompletos = {
      ...datos,
      folio: folio,
      tecnico: sesion.nombre || sesion.email || 'Desconocido',
      tecnicoEmail: sesion.email,
      fechaReporte: ahora.toISOString(),
      fechaActualizacion: ahora.toISOString(),
      esNuevo: esNuevo,
      folderUrl: folderUrl
    };

    let estadoActual = datos.estado || 'BORRADOR';
    if (esNuevo && !datos.estado) {
      estadoActual = 'BORRADOR';
    } else if (!esNuevo && datos.estado) {
      estadoActual = datos.estado;
    }

    const datosJSON = JSON.stringify(datosCompletos);

    // 7. Guardar en Bitácora
    if (esNuevo) {
      const nuevaFila = sheet.getLastRow() + 1;
      const filaData = new Array(headers.length).fill('');
      
      if (idxFolio !== -1) filaData[idxFolio] = folio;
      if (idxEstado !== -1) filaData[idxEstado] = estadoActual;
      if (idxTecnico !== -1) filaData[idxTecnico] = sesion.nombre || sesion.email || '';
      if (idxFecha !== -1) filaData[idxFecha] = ahora;
      if (idxDatos !== -1) filaData[idxDatos] = datosJSON;
      if (idxFotos !== -1) filaData[idxFotos] = folderUrl;
      if (idxGPS !== -1 && datos.gps) filaData[idxGPS] = datos.gps;
      
      sheet.getRange(nuevaFila, 1, 1, filaData.length).setValues([filaData]);
      console.log(`✅ Nuevo reporte creado: ${folio}`);
    } else {
      if (filaDestino === -1) {
        throw new Error(`No se encontró el folio ${folio} para actualizar`);
      }
      
      if (idxEstado !== -1) {
        sheet.getRange(filaDestino, idxEstado + 1).setValue(estadoActual);
      }
      if (idxDatos !== -1) {
        let datosExistentes = {};
        try {
          const datosViejos = registroExistente ? registroExistente[idxDatos] : null;
          if (datosViejos) {
            datosExistentes = typeof datosViejos === 'string' ? JSON.parse(datosViejos) : datosViejos;
          }
        } catch (e) {
          console.warn('⚠️ Error al parsear datos existentes:', e);
        }
        const datosActualizados = { ...datosExistentes, ...datosCompletos };
        sheet.getRange(filaDestino, idxDatos + 1).setValue(JSON.stringify(datosActualizados));
      }
      if (idxFotos !== -1 && folderUrl) {
        sheet.getRange(filaDestino, idxFotos + 1).setValue(folderUrl);
      }
      if (idxFecha !== -1) {
        sheet.getRange(filaDestino, idxFecha + 1).setValue(ahora);
      }
      if (idxGPS !== -1 && datos.gps) {
        sheet.getRange(filaDestino, idxGPS + 1).setValue(datos.gps);
      }
      if (idxTecnico !== -1 && sesion.nombre) {
        sheet.getRange(filaDestino, idxTecnico + 1).setValue(sesion.nombre);
      }
      
      console.log(`✅ Reporte actualizado: ${folio}`);
    }

    // 8. Retornar resultado
    return {
      ok: true,
      folio: folio,
      folderUrl: folderUrl || '',
      esNuevo: esNuevo,
      estado: estadoActual,
      mensaje: esNuevo 
        ? `Reporte ${folio} creado exitosamente` 
        : `Reporte ${folio} actualizado exitosamente`,
      timestamp: ahora.toISOString()
    };

  } catch (err) {
    console.error('❌ Error en recibirReporte:', err);
    return {
      ok: false,
      error: 'Error al procesar el reporte: ' + err.message,
      folio: datos?.folio || null,
      timestamp: new Date().toISOString()
    };
  }
}
// ============================================================
// 🔹 FUNCIÓN AUXILIAR QUE FALTABA
// ============================================================

/**
 * Anexa fotos nuevas a una carpeta de Drive existente
 * @param {string} folderUrl - URL de la carpeta existente en Drive
 * @param {Array} archivos - Arreglo de archivos (fotos) en Base64
 * @param {Object} opciones - Opciones adicionales (opcional)
 * @returns {Object} { success: boolean, folderUrl: string, archivosSubidos: number, errores: Array }
 */
function _anexarFotosACarpetaExistente(folderUrl, archivos, opciones = {}) {
  const resultado = {
    success: false,
    folderUrl: folderUrl,
    archivosSubidos: 0,
    errores: []
  };

  try {
    // 1. Validaciones de entrada
    if (!archivos || !Array.isArray(archivos) || archivos.length === 0) {
      resultado.errores.push('No hay archivos para subir');
      console.warn('⚠️ _anexarFotosACarpetaExistente: No hay archivos para subir');
      return resultado;
    }

    if (!folderUrl || folderUrl === '' || folderUrl === 'root') {
      resultado.errores.push('URL de carpeta inválida');
      console.warn('⚠️ _anexarFotosACarpetaExistente: URL de carpeta inválida');
      return resultado;
    }

    // 2. Extraer ID de la carpeta
    let folderId = null;
    const patrones = [
      /[-\w]{25,}/,
      /folders\/([-\w]{25,})/,
      /file\/d\/([-\w]{25,})/,
      /id=([-\w]{25,})/
    ];

    for (const patron of patrones) {
      const match = folderUrl.match(patron);
      if (match) {
        folderId = match[1] || match[0];
        break;
      }
    }

    if (!folderId) {
      try {
        const carpeta = DriveApp.getFolderById(folderUrl);
        folderId = carpeta.getId();
      } catch (err) {
        const carpetas = DriveApp.getFoldersByName(folderUrl);
        if (carpetas.hasNext()) {
          const carpeta = carpetas.next();
          folderId = carpeta.getId();
        } else {
          resultado.errores.push(`No se pudo identificar la carpeta: ${folderUrl}`);
          console.error('❌ _anexarFotosACarpetaExistente: No se pudo identificar la carpeta');
          return resultado;
        }
      }
    }

    // 3. Obtener la carpeta
    let carpeta;
    try {
      carpeta = DriveApp.getFolderById(folderId);
      console.log(`✅ Carpeta encontrada: ${carpeta.getName()} (ID: ${folderId})`);
    } catch (err) {
      resultado.errores.push(`Error al acceder a la carpeta: ${err.message}`);
      console.error('❌ _anexarFotosACarpetaExistente: Error al acceder a la carpeta:', err.message);
      return resultado;
    }

    // 4. Crear subcarpeta por fecha (opcional)
    let subcarpeta = null;
    if (opciones.organizarPorFecha !== false) {
      try {
        const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const nombreSubcarpeta = `Fotos_anexadas_${fechaStr}`;
        const subcarpetas = carpeta.getFoldersByName(nombreSubcarpeta);
        if (subcarpetas.hasNext()) {
          subcarpeta = subcarpetas.next();
        } else {
          subcarpeta = carpeta.createFolder(nombreSubcarpeta);
        }
        console.log(`📁 Subcarpeta creada/encontrada: ${subcarpeta.getName()}`);
      } catch (err) {
        console.warn('⚠️ No se pudo crear subcarpeta, se subirán a la raíz:', err.message);
      }
    }

    const carpetaDestino = subcarpeta || carpeta;
    const archivosSubidos = [];
    const erroresSubida = [];

    // 5. Subir cada archivo
    archivos.forEach((archivo, index) => {
      try {
        if (!archivo) {
          erroresSubida.push(`Archivo ${index}: Datos vacíos`);
          return;
        }

        const nombreArchivo = archivo.nombre || archivo.name || `foto_${Date.now()}_${index}.jpg`;
        const base64Data = archivo.base64 || archivo.data || archivo;
        const mimeType = archivo.mimeType || archivo.mime || 'image/jpeg';

        if (!base64Data || typeof base64Data !== 'string' || base64Data.length === 0) {
          erroresSubida.push(`Archivo ${index} (${nombreArchivo}): Datos Base64 inválidos`);
          return;
        }

        let blob;
        try {
          let cleanBase64 = base64Data;
          if (cleanBase64.includes('base64,')) {
            cleanBase64 = cleanBase64.split('base64,')[1];
          }
          const bytes = Utilities.base64Decode(cleanBase64);
          blob = Utilities.newBlob(bytes, mimeType, nombreArchivo);
        } catch (err) {
          erroresSubida.push(`Archivo ${index} (${nombreArchivo}): Error al decodificar: ${err.message}`);
          return;
        }

        const file = carpetaDestino.createFile(blob);
        archivosSubidos.push({
          nombre: file.getName(),
          id: file.getId(),
          url: file.getUrl(),
          tamaño: file.getSize()
        });

        console.log(`✅ Archivo subido: ${file.getName()} (${(file.getSize() / 1024).toFixed(2)} KB)`);

      } catch (err) {
        const nombreError = archivo?.nombre || archivo?.name || `Archivo ${index}`;
        erroresSubida.push(`Error al subir ${nombreError}: ${err.message}`);
        console.error(`❌ Error al subir archivo ${index}:`, err.message);
      }
    });

    // 6. Actualizar resultado
    resultado.success = archivosSubidos.length > 0;
    resultado.archivosSubidos = archivosSubidos.length;
    resultado.folderUrl = carpetaDestino.getUrl();
    resultado.archivosDetalle = archivosSubidos;
    resultado.errores = erroresSubida;

    if (erroresSubida.length > 0) {
      console.warn(`⚠️ ${erroresSubida.length} errores en la subida de archivos:`, erroresSubida);
    }

    console.log(`✅ Proceso completado: ${archivosSubidos.length} archivos subidos a ${carpetaDestino.getName()}`);
    return resultado;

  } catch (err) {
    console.error('❌ _anexarFotosACarpetaExistente: Error crítico:', err.message);
    resultado.success = false;
    resultado.errores.push(`Error crítico: ${err.message}`);
    return resultado;
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
  let folderId = params['DRIVE_FOLDER_ID'];
  
  // ✅ VALIDACIÓN: Si no hay carpeta configurada, crearla
  if (!folderId || folderId === '' || folderId === 'root') {
    // Crear carpeta principal con nombre descriptivo
    const nombreCarpeta = 'Reportes_Fleet_Manager_' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'
    );
    const nuevaCarpeta = DriveApp.createFolder(nombreCarpeta);
    folderId = nuevaCarpeta.getId();
    
    // Guardar el ID en parámetros para futuros usos
    _escribirParam('DRIVE_FOLDER_ID', folderId);
    console.log('📁 Carpeta principal creada:', nombreCarpeta, 'ID:', folderId);
  }

  // ✅ VALIDACIÓN: Intentar obtener la carpeta, si falla crearla
  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(folderId);
  } catch (err) {
    console.warn('⚠️ No se pudo obtener la carpeta por ID, creando una nueva...', err.message);
    // Crear nueva carpeta como fallback
    const nombreCarpeta = 'Reportes_Fleet_Manager_' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'
    );
    rootFolder = DriveApp.createFolder(nombreCarpeta);
    folderId = rootFolder.getId();
    _escribirParam('DRIVE_FOLDER_ID', folderId);
    console.log('📁 Carpeta principal recreada:', nombreCarpeta, 'ID:', folderId);
  }

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

// ════════════════════════════════════════════════════════════
// 6. FUNCIONES DEL REVISOR (Nivel 2)
// ════════════════════════════════════════════════════════════

/**
 * Obtiene registros de la bitácora con paginación y filtros
 * @param {string} token - Token de sesión
 * @param {Object} filtros - { estado, tecnico, fechaDesde, fechaHasta, pagina, limite }
 * @returns {Object} { ok: true, registros: [], total: number, pagina: number, limite: number }
 */
function obtenerRegistros(token, filtros) {
  try {
    console.log('📊 obtenerRegistros - INICIO', new Date().toISOString());

    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    var sesion = sesionResp.sesion;
    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Bitácora.' };
    }

    // ✅ PARÁMETROS DE PAGINACIÓN
    var pagina = parseInt(filtros.pagina) || 1;
    var limite = parseInt(filtros.limite) || 25;
    var inicio = (pagina - 1) * limite;

    var datos = sheet.getDataRange().getValues();
    var headers = datos[0];

    // ✅ FILTRAR Y CONTAR REGISTROS
    var registrosFiltrados = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila || !fila[0]) continue;

      // ✅ FILTROS BÁSICOS
      if (filtros.estado && fila[12] !== filtros.estado) continue;
      if (filtros.tecnico && fila[7] !== filtros.tecnico) continue;

      // ✅ FILTRO DE FECHAS
      if (filtros.fechaDesde) {
        var fechaDesde = new Date(filtros.fechaDesde);
        var fechaServicio = new Date(fila[1]);
        if (fechaServicio < fechaDesde) continue;
      }
      if (filtros.fechaHasta) {
        var fechaHasta = new Date(filtros.fechaHasta);
        var fechaServicio = new Date(fila[1]);
        if (fechaServicio > fechaHasta) continue;
      }

      // ✅ SI ES TÉCNICO, SOLO SUS REGISTROS
      if (sesion.rol === 3 && fila[7]?.toString() !== sesion.usuarioId) continue;

      // ✅ CONSTRUIR REGISTRO
      var reg = {};
      for (var j = 0; j < headers.length; j++) {
        var valor = fila[j];
        if (valor instanceof Date) {
          valor = valor.toISOString();
        } else if (valor === undefined || valor === null) {
          valor = '';
        } else if (typeof valor === 'boolean' || typeof valor === 'number') {
          // dejar tal cual
        } else {
          valor = valor.toString();
        }
        reg[headers[j]] = valor;
      }
      registrosFiltrados.push(reg);
    }

    // ✅ CALCULAR TOTALES
    var total = registrosFiltrados.length;
    var totalPaginas = Math.ceil(total / limite);

    // ✅ OBTENER PÁGINA SOLICITADA
    var registrosPagina = registrosFiltrados.slice(inicio, inicio + limite);

    console.log('📊 Registros:', {
      total: total,
      pagina: pagina,
      limite: limite,
      totalPaginas: totalPaginas,
      mostrando: registrosPagina.length
    });

    return {
      ok: true,
      registros: registrosPagina,
      total: total,
      pagina: pagina,
      limite: limite,
      totalPaginas: totalPaginas,
      rol: sesion.rol
    };

  } catch (err) {
    console.error('❌ Error en obtenerRegistros:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Crea una vista materializada de la Bitácora para consultas rápidas
 * Ejecutar una vez al día o con un trigger
 */
function crearVistaBitacora() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bitacora = ss.getSheetByName('📝_Bitacora_Revisiones');
  var vista = ss.getSheetByName('📊_Vista_Bitacora');
  
  if (!vista) {
    vista = ss.insertSheet('📊_Vista_Bitacora');
  }
  
  // Copiar solo las columnas necesarias
  var datos = bitacora.getDataRange().getValues();
  vista.clear();
  vista.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
  
  // Agregar filtros y formato
  vista.getRange(1, 1, 1, datos[0].length).setFontWeight('bold');
  
  console.log('✅ Vista creada con ' + datos.length + ' registros');
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
 * Genera notificación al técnico que creó el reporte.
 */
function aprobarRegistro(token, folio) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  if (sesion.rol > 2) return { ok: false, error: 'Sin permisos para aprobar.' };

  var sheet = SHEETS.BITACORA();
  var datos = sheet.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === folio) {
      var fila = i + 1;
      var tecnicoId = (datos[i][7] || '').toString().trim(); // TECNICO_ID

      // Actualizar estado
      sheet.getRange(fila, 13).setValue('Aprobado');      // M: ESTADO
      sheet.getRange(fila, 15).setValue(sesion.nombre);    // O: APROBADO_POR
      sheet.getRange(fila, 16).setValue(new Date());       // P: FECHA_APROBACION

      // 🔔 NOTIFICACIÓN: Reporte aprobado (al técnico)
      if (tecnicoId) {
        _crearNotificacion(
          tecnicoId,
          'APROBACION',
          '✅ Tu reporte ' + folio + ' ha sido aprobado por ' + sesion.nombre,
          folio,
          '#panel-mis-registros'
        );
      }

      // 🔔 NOTIFICACIÓN: Reporte aprobado (a Admin/Revisores)
      var admins = _obtenerUsuariosPorRol([1, 2]);
      for (var u = 0; u < admins.length; u++) {
        if (admins[u] !== sesion.usuarioId) {
          _crearNotificacion(
            admins[u],
            'APROBACION',
            '✅ El reporte ' + folio + ' fue aprobado por ' + sesion.nombre,
            folio,
            '#panel-registros'
          );
        }
      }

      // ✅ AUDITORÍA: Reporte aprobado
      _registrarAuditoria(
        sesion.usuarioId,
        sesion.nombre,
        'APROBAR_REPORTE',
        'REPORTES',
        'Reporte ' + folio + ' aprobado por ' + sesion.nombre + ' (Técnico: ' + (datos[i][8] || 'N/A') + ')',
        folio,
        '',
        ''
      );

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

  if (sesion.rol > 3) return { ok: false, error: 'Sin permisos para reportar.' };

  try {
    var ahora = new Date();
    var latitud = ubicacion ? ubicacion.latitud : null;
    var longitud = ubicacion ? ubicacion.longitud : null;

    var sheet = SHEETS.BITACORA();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Bitácora' };
    }

    // ✅ GENERAR FOLIO
    var folio = generarFolio();

    // ✅ SUBIR FOTOS
    var folderUrl = '';
    if (archivos && archivos.length > 0) {
      try {
        folderUrl = _subirFotosDrive(folio, archivos);
      } catch (fotoError) {
        console.warn('⚠️ Error al subir fotos:', fotoError.message);
        folderUrl = '';
      }
    }

    // ✅ PREPARAR DATOS PARA BITÁCORA
    // El tipo de servicio se deja como "Pendiente" o vacío para que el Revisor lo asigne
    var headers = sheet.getDataRange().getValues()[0];
    var filaData = new Array(headers.length).fill('');

    // Mapear índices
    var idxFolio = headers.indexOf('FOLIO');
    var idxFechaReporte = headers.indexOf('FECHA_REPORTE');
    var idxTecnico = headers.indexOf('TECNICO');
    var idxGPS = headers.indexOf('GPS');
    var idxCliente = headers.indexOf('CLIENTE');
    var idxTipoRevision = headers.indexOf('TIPO_REVISION');
    var idxObservaciones = headers.indexOf('OBSERVACIONES');
    var idxEstado = headers.indexOf('ESTADO');
    var idxMontoTotal = headers.indexOf('MONTO_TOTAL');
    var idxFotos = headers.indexOf('FOTOS_DRIVE_URL');
    var idxDatosJSON = headers.indexOf('DATOS_JSON');
    var idxEstadoPago = headers.indexOf('ESTADO_PAGO');

    // Llenar datos
    if (idxFolio !== -1) filaData[idxFolio] = folio;
    if (idxFechaReporte !== -1) filaData[idxFechaReporte] = ahora;
    if (idxTecnico !== -1) filaData[idxTecnico] = sesion.nombre || sesion.email;
    if (idxGPS !== -1) filaData[idxGPS] = datos.gps || datos.economico || '';
    if (idxCliente !== -1) filaData[idxCliente] = datos.cliente || '';
    
    // ✅ TIPO_REVISION queda como "Pendiente" para que el Revisor lo asigne
    if (idxTipoRevision !== -1) filaData[idxTipoRevision] = 'Pendiente';
    
    if (idxObservaciones !== -1) filaData[idxObservaciones] = datos.detalleTrabajo || '';
    
    // ✅ ESTADO: "Borrador" o "Listo para pago" según el caso
    var estado = datos.esBorrador ? 'Borrador' : 'Listo para pago';
    if (idxEstado !== -1) filaData[idxEstado] = estado;
    
    if (idxMontoTotal !== -1) filaData[idxMontoTotal] = 0; // El Revisor asignará el precio
    if (idxFotos !== -1) filaData[idxFotos] = folderUrl;
    if (idxDatosJSON !== -1) filaData[idxDatosJSON] = JSON.stringify(datos);
    if (idxEstadoPago !== -1) filaData[idxEstadoPago] = 'Pendiente';

    // Guardar en Bitácora
    var nuevaFila = sheet.getLastRow() + 1;
    sheet.getRange(nuevaFila, 1, 1, filaData.length).setValues([filaData]);

    console.log('📋 Reporte creado:', folio, 'Estado:', estado);

    // ✅ NOTIFICACIÓN AL REVISOR
    _crearNotificacion(
      'REVISOR', // o a todos los revisores
      'NUEVO_REPORTE',
      '📝 Nuevo reporte ' + folio + ' de ' + sesion.nombre + ' espera revisión.',
      folio,
      '#panel-registros'
    );

    return {
      ok: true,
      folio: folio,
      folderUrl: folderUrl,
      mensaje: 'Reporte creado exitosamente',
      estado: estado
    };

  } catch (err) {
    console.error('❌ Error en recibirReporteMejorado:', err);
    return { ok: false, error: 'Error al procesar el reporte: ' + err.message };
  }
}

/**
 * Obtiene la serie GPS instalada en un vehículo (la primera que encuentra)
 * @param {string} economico - ID del vehículo
 * @returns {string|null} - Serie GPS o null si no tiene
 */
function _obtenerSerieGPSporEconomico(economico) {
  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) return null;

    var datos = sheet.getDataRange().getValues();
    var economicoStr = economico.toString().toUpperCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var serie = datos[i][0];
      if (!serie) continue;

      var tipo = (datos[i][1] || '').toString().toUpperCase();
      var estado = (datos[i][4] || '').toString().toUpperCase().trim();
      var economicoAsignado = (datos[i][5] || '').toString().toUpperCase().trim();

      if (economicoAsignado === economicoStr && estado === 'INSTALADO') {
        // Buscar solo Gateways (VG)
        if (tipo.indexOf('VG') !== -1 || tipo.indexOf('GATEWAY') !== -1) {
          return serie.toString().trim();
        }
      }
    }

    return null;

  } catch (err) {
    console.error('Error en _obtenerSerieGPSporEconomico:', err);
    return null;
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
  try {
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      console.warn('No se encontró la hoja de inventario');
      return;
    }

    var datos = sheet.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString() === serieGPS.toString()) {
        // Cambiar estado a "Disponible"
        sheet.getRange(i + 1, 5).setValue('Disponible');
        // Limpiar económico asignado
        sheet.getRange(i + 1, 6).setValue('');
        // Limpiar fecha de instalación
        sheet.getRange(i + 1, 7).setValue('');
        // Actualizar última modificación
        sheet.getRange(i + 1, 9).setValue(new Date());
        console.log('✅ Serie liberada:', serieGPS);
        break;
      }
    }
  } catch (err) {
    console.error('Error en _liberarSerieGPS:', err);
  }
}


/**
 * Actualiza el estado del GPS en Inventario al ser instalado
 * Busca la serie ignorando guiones para máxima compatibilidad
 * @param {string} serieGPS - Serie del GPS a instalar
 * @param {string} economico - Número económico del vehículo
 * @returns {Object} { ok: boolean, error: string|null, mensaje: string }
 */
function _actualizarEstadoGPS(serieGPS, economico) {
  try {
    // ============================================================
    // 1. VALIDAR ENTRADA
    // ============================================================
    
    var sheet = SHEETS.INVENTARIO();
    if (!sheet) {
      console.warn('❌ No se encontró la hoja INVENTARIO');
      return { ok: false, error: 'No se encontró la hoja INVENTARIO.' };
    }

    if (!serieGPS || serieGPS.toString().trim() === '') {
      console.warn('⚠️ Serie GPS vacía, no se puede actualizar');
      return { ok: false, error: 'La serie GPS provista está vacía.' };
    }

    // ============================================================
    // 2. LIMPIAR SERIE PARA BÚSQUEDA
    // ============================================================
    
    var serieBusquedaClean = serieGPS.toString()
      .toUpperCase()
      .replace(/[-_\s]/g, '')
      .trim();
    
    var economicoStr = (economico || '').toString().trim();

    console.log('🔍 Buscando serie limpia:', serieBusquedaClean);
    console.log('📌 Para económico:', economicoStr);

    // ============================================================
    // 3. BUSCAR SERIE EN INVENTARIO
    // ============================================================
    
    var datos = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaEncontrada = -1;
    var serieOriginal = '';

    for (var i = 1; i < datos.length; i++) {
      var serieCeldaRaw = datos[i][0];
      if (!serieCeldaRaw) continue;

      // Limpiar la serie del inventario de la misma forma
      var serieCeldaClean = serieCeldaRaw.toString()
        .toUpperCase()
        .replace(/[-_\s]/g, '')
        .trim();

      // Comparación sin guiones
      if (serieCeldaClean === serieBusquedaClean) {
        filaEncontrada = i + 1;
        encontrado = true;
        serieOriginal = serieCeldaRaw.toString();
        console.log('✅ Coincidencia encontrada en fila:', filaEncontrada);
        console.log('   Serie en inventario:', serieOriginal);
        break;
      }
    }

    if (!encontrado) {
      console.warn('⚠️ NO se encontró la serie en inventario:', serieGPS);
      console.warn('   Buscando como:', serieBusquedaClean);
      console.warn('   Revisa que la serie esté registrada en 📦_Inventario_GPS');
      return { 
        ok: false, 
        error: 'La serie no fue hallada en el inventario.',
        serieBuscada: serieBusquedaClean
      };
    }

    // ============================================================
    // 4. ACTUALIZAR REGISTRO
    // ============================================================
    
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

    return { 
      ok: true, 
      error: null,
      mensaje: 'GPS instalado correctamente',
      fila: filaEncontrada
    };

  } catch (err) {
    console.error('❌ Error en _actualizarEstadoGPS:', err);
    return { 
      ok: false, 
      error: 'Error al actualizar GPS: ' + err.message 
    };
  }
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
 * Obtiene los dispositivos instalados en un vehículo
 * @param {string} token - Token de sesión
 * @param {string} economico - Número económico del vehículo
 * @returns {Object} { ok, gateway, camara, accesorios }
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

        // ✅ CAMBIO 1: Incluir "GARANTÍA" junto con "INSTALADO"
        if (economicoAsignadoStr === economicoStr && (estadoStr === 'INSTALADO' || estadoStr === 'GARANTÍA')) {
          var tipoStr = tipo.toString().toUpperCase();

          if (tipoStr.indexOf('GATEWAY') !== -1 || tipoStr.indexOf('VG') !== -1) {
            resultado.gateway = {
              serie: (serie || '').toString(),
              tipo: tipo.toString(),
              estado: estado.toString() // ✅ CAMBIO 2: Añadir estado para que el frontend lo muestre
            };
            console.log('✅ Gateway encontrado:', resultado.gateway);
          } else if (tipoStr.indexOf('CAMARA') !== -1 || tipoStr.indexOf('CM') !== -1) {
            resultado.camara = {
              serie: (serie || '').toString(),
              tipo: tipo.toString(),
              estado: estado.toString() // ✅ CAMBIO 2: Añadir estado para que el frontend lo muestre
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
 * NOTA: La columna ECONOMICO_ASIGNADO en inventario referencia a ECONOMICO en catálogo
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
        var eco = (catalogoData[c][0] || '').toString().toUpperCase().trim(); // Columna 1: ECONOMICO
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
          equiposPorEconomico[economico].gateway = tipo;        // ← Guarda el TIPO
          equiposPorEconomico[economico].gatewaySerie = serie;  // ← Guarda la SERIE
        } else if (tipoUpper.indexOf('CM') !== -1 || tipoUpper.indexOf('CAMARA') !== -1) {
          equiposPorEconomico[economico].camara = tipo;         // ← Guarda el TIPO
          equiposPorEconomico[economico].camaraSerie = serie;   // ← Guarda la SERIE
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
          var listaEconomicos = economicoAsignado.split(',').map(function (e) { return e.trim(); });
          listaEconomicos.forEach(function (eco) {
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
 * Genera un ID único para tickets
 */
function generarIdTicket() {
  var params = _leerParams();
  var prefijo = params['TICKET_PREFIJO'] || 'TKT';
  var ultimo = parseInt(params['TICKET_ULTIMO'] || '0', 10);
  var nuevo = ultimo + 1;
  _escribirParam('TICKET_ULTIMO', String(nuevo));
  return prefijo + '-' + String(nuevo).padStart(4, '0');
}


// ============================================================
// TICKETS - BACKEND
// ============================================================


/**
 * Obtiene la lista de tickets con filtros y control de visibilidad por técnico
 * @param {string} token - Token de sesión
 * @param {Object} filtros - { estado, unidad, tecnico }
 * @returns {Object} { ok: boolean, tickets: [], total: number }
 */
function obtenerTickets(token, filtros) {
  try {
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      return { ok: false, error: sesionResp.error };
    }

    const sesion = sesionResp.sesion;
    const sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Tickets' };
    }

    const datos = sheet.getDataRange().getValues();
    if (datos.length < 2) {
      return { ok: true, tickets: [], total: 0 };
    }

    const headers = datos[0];

    // Mapear índices de columnas
    const idxId = headers.indexOf('ID');
    const idxFecha = headers.indexOf('FECHA');
    const idxUnidad = headers.indexOf('UNIDAD');
    const idxDescripcion = headers.indexOf('DESCRIPCION');
    const idxCreadoPor = headers.indexOf('CREADO_POR');
    const idxCreadoPorNombre = headers.indexOf('CREADO_POR_NOMBRE');
    const idxEstado = headers.indexOf('ESTADO');
    const idxTecnicoAsignado = headers.indexOf('TECNICO_ASIGNADO');
    const idxTecnicoNombre = headers.indexOf('TECNICO_NOMBRE');
    const idxFechaCierre = headers.indexOf('FECHA_CIERRE');
    const idxComentarios = headers.indexOf('COMENTARIOS');
    const idxUltimaActualizacion = headers.indexOf('ULTIMA_ACTUALIZACION');
    const idxTecnicosAutorizados = headers.indexOf('TECNICOS_AUTORIZADOS');

    // Verificar columnas mínimas necesarias
    if (idxId === -1) {
      return { ok: false, error: 'Estructura de tickets incompleta: falta columna ID' };
    }

    const tickets = [];
    const usuarioId = sesion.usuarioId;
    const esAdmin = sesion.rol === 1 || sesion.rol === 2; // Admin o Revisor

    for (var i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila || !fila[idxId]) continue;

      // ============================================================
      // 1. CONTROL DE VISIBILIDAD
      // ============================================================
      let puedeVer = false;

      // Admin/Revisor: ven todos los tickets
      if (esAdmin) {
        puedeVer = true;
      }
      // Creador del ticket: puede verlo
      else if (fila[idxCreadoPor] === usuarioId) {
        puedeVer = true;
      }
      // Técnico autorizado: puede verlo
      else if (idxTecnicosAutorizados !== -1) {
        const autorizados = fila[idxTecnicosAutorizados]
          ? fila[idxTecnicosAutorizados].toString().split(',')
          : [];
        puedeVer = autorizados.includes(usuarioId);
      }
      // Técnico asignado: puede verlo
      else if (fila[idxTecnicoAsignado] === usuarioId) {
        puedeVer = true;
      }

      if (!puedeVer) continue;

      // ============================================================
      // 2. APLICAR FILTROS
      // ============================================================
      if (filtros) {
        if (filtros.estado && fila[idxEstado] !== filtros.estado) continue;
        if (filtros.unidad && fila[idxUnidad] && 
            !fila[idxUnidad].toLowerCase().includes(filtros.unidad.toLowerCase())) continue;
        if (filtros.tecnico && fila[idxTecnicoNombre] && 
            !fila[idxTecnicoNombre].toLowerCase().includes(filtros.tecnico.toLowerCase())) continue;
      }

      // ============================================================
      // 3. CONSTRUIR OBJETO TICKET
      // ============================================================
      const ticket = {
        id: fila[idxId] || '',
        fecha: fila[idxFecha] ? _formatearFecha(fila[idxFecha]) : '—',
        unidad: fila[idxUnidad] || '—',
        descripcion: fila[idxDescripcion] || '—',
        creadoPor: fila[idxCreadoPor] || '',
        creadoPorNombre: fila[idxCreadoPorNombre] || fila[idxCreadoPor] || '—',
        estado: fila[idxEstado] || 'Pendiente',
        tecnico: fila[idxTecnicoAsignado] || '',
        tecnicoNombre: fila[idxTecnicoNombre] || '—',
        fechaCierre: fila[idxFechaCierre] ? _formatearFecha(fila[idxFechaCierre]) : null,
        comentarios: fila[idxComentarios] || '',
        ultimaActualizacion: fila[idxUltimaActualizacion] ? _formatearFecha(fila[idxUltimaActualizacion]) : '—',
        tecnicosAutorizados: fila[idxTecnicosAutorizados] || ''
      };

      tickets.push(ticket);
    }

    // ============================================================
    // 4. ORDENAR POR FECHA (más recientes primero)
    // ============================================================
    tickets.sort(function(a, b) {
      // Intentar ordenar por fecha real (si están en formato ISO o Date)
      const fechaA = new Date(a.fecha);
      const fechaB = new Date(b.fecha);
      if (!isNaN(fechaA) && !isNaN(fechaB)) {
        return fechaB - fechaA;
      }
      return 0;
    });

    return {
      ok: true,
      tickets: tickets,
      total: tickets.length
    };

  } catch (err) {
    console.error('❌ Error en obtenerTickets:', err);
    return { ok: false, error: 'Error al obtener tickets: ' + err.message };
  }
}
/**
 * Actualiza la lista de técnicos autorizados para un ticket
 * @param {string} token - Token de sesión (debe ser admin)
 * @param {string} ticketId - ID del ticket
 * @param {string[]} tecnicosIds - Lista de IDs de técnicos autorizados
 * @returns {Object} { ok: boolean, error: string }
 */
function actualizarVisibilidadTicket(token, ticketId, tecnicosIds) {
  try {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
    if (sesionResp.sesion.rol !== 1) {
      return { ok: false, error: 'Solo administradores pueden modificar visibilidad.' };
    }

    var sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de tickets.' };
    }

    var datos = sheet.getDataRange().getValues();
    var headers = datos[0];
    var idxId = headers.indexOf('ID');
    var idxAutorizados = headers.indexOf('TECNICOS_AUTORIZADOS');

    if (idxId === -1 || idxAutorizados === -1) {
      return { ok: false, error: 'Estructura de tickets incompleta.' };
    }

    var filaEncontrada = -1;
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxId] === ticketId) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { ok: false, error: 'Ticket no encontrado.' };
    }

    var lista = tecnicosIds && tecnicosIds.length > 0 ? tecnicosIds.join(',') : '';
    sheet.getRange(filaEncontrada, idxAutorizados + 1).setValue(lista);

    _registrarAuditoria(
      sesionResp.sesion.usuarioId,
      sesionResp.sesion.nombre,
      'ACTUALIZAR_VISIBILIDAD_TICKET',
      'TICKETS',
      'Visibilidad del ticket ' + ticketId + ' actualizada',
      ticketId,
      '',
      ''
    );

    return { ok: true };

  } catch (err) {
    console.error('Error en actualizarVisibilidadTicket:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Obtiene la lista de técnicos activos para checkboxes
 */
function obtenerTecnicosParaTicket(token) {
  try {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

    var sheet = SHEETS.USUARIOS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de usuarios.' };
    }

    var datos = sheet.getDataRange().getValues();
    var tecnicos = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var id = fila[0];
      var nombre = fila[1];
      var rol = fila[4];
      var activo = fila[5];

      if (id && activo === true && Number(rol) === 3) {
        tecnicos.push({ id: id.toString(), nombre: nombre.toString() });
      }
    }

    return { ok: true, tecnicos: tecnicos };

  } catch (err) {
    return { ok: false, error: err.message };
  }
}
/**
 * Obtiene un ticket específico para edición
 * @param {string} token - Token de sesión
 * @param {string} ticketId - ID del ticket
 * @returns {Object} { ok, ticket }
 */
function obtenerTicket(token, ticketId) {
  try {
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      return { ok: false, error: sesionResp.error };
    }

    const sesion = sesionResp.sesion;
    const sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Tickets' };
    }

    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];

    // Mapear índices
    const idxId = headers.indexOf('ID');
    const idxUnidad = headers.indexOf('UNIDAD');
    const idxDescripcion = headers.indexOf('DESCRIPCION');
    const idxEstado = headers.indexOf('ESTADO');
    const idxPrioridad = headers.indexOf('PRIORIDAD');
    const idxComentarios = headers.indexOf('COMENTARIOS');
    const idxNotasInternas = headers.indexOf('NOTAS_INTERNAS');
    const idxCreadoPor = headers.indexOf('CREADO_POR');

    if (idxId === -1) {
      return { ok: false, error: 'Estructura de tickets incompleta: falta columna ID' };
    }

    const esAdmin = sesion.rol === 1 || sesion.rol === 2;

    // Buscar el ticket
    let fila = null;
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxId] === ticketId) {
        fila = datos[i];
        break;
      }
    }

    if (!fila) {
      return { ok: false, error: 'Ticket no encontrado: ' + ticketId };
    }

    // ✅ Verificar permisos (Admin/Revisor o creador del ticket)
    if (!esAdmin && fila[idxCreadoPor] !== sesion.usuarioId) {
      return { ok: false, error: 'No tienes permisos para editar este ticket.' };
    }

    // Construir objeto
    const ticket = {
      ID: fila[idxId] || '',
      UNIDAD: fila[idxUnidad] || '',
      DESCRIPCION: fila[idxDescripcion] || '',
      ESTADO: fila[idxEstado] || 'Pendiente',
      PRIORIDAD: fila[idxPrioridad] || 'Media',
      COMENTARIOS: fila[idxComentarios] || '',
      NOTAS_INTERNAS: fila[idxNotasInternas] || ''
    };

    return { ok: true, ticket: ticket };

  } catch (err) {
    console.error('❌ Error en obtenerTicket:', err);
    return { ok: false, error: err.message };
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
 * Crea un nuevo ticket con control de visibilidad por técnicos
 * @param {string} token - Token de sesión
 * @param {Object} datos - Datos del ticket
 * @returns {Object} { ok: boolean, id: string, error: string }
 */
function crearTicket(token, datos) {
  try {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
    var sesion = sesionResp.sesion;

    if (!datos.unidad || !datos.descripcion) {
      return { ok: false, error: 'Unidad y descripción son obligatorios.' };
    }

    // Generar ID
    var id = generarIdTicket();

    // Preparar técnicos autorizados
    var tecnicosAutorizados = [];
    if (sesion.rol === 1) {
      tecnicosAutorizados = datos.tecnicosAutorizados || [];
    } else {
      tecnicosAutorizados = [sesion.usuarioId];
    }

    var sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de tickets.' };
    }

    var ahora = new Date();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Mapear columnas
    var colIndex = function(nombre) { return headers.indexOf(nombre) + 1; };

    var fila = [
      id,                          // ID
      ahora,                       // FECHA
      datos.unidad,                // UNIDAD
      datos.descripcion,           // DESCRIPCION
      sesion.usuarioId,            // CREADO_POR
      sesion.nombre,               // CREADO_POR_NOMBRE
      'Pendiente',                 // ESTADO
      '',                          // TECNICO_ASIGNADO
      '',                          // TECNICO_NOMBRE
      '',                          // FECHA_CIERRE
      '',                          // COMENTARIOS
      ahora,                       // ULTIMA_ACTUALIZACION
      tecnicosAutorizados.join(','), // TECNICOS_AUTORIZADOS
      sesion.rol,                  // CREADO_POR_ROL
      datos.prioridad || 'Media',  // PRIORIDAD
      datos.categoria || 'Otro',   // CATEGORIA
      '',                          // ARCHIVOS_ADJUNTOS
      datos.notasInternas || ''    // NOTAS_INTERNAS
    ];

    sheet.appendRow(fila);

    _registrarAuditoria(
      sesion.usuarioId,
      sesion.nombre,
      'CREAR_TICKET',
      'TICKETS',
      'Ticket ' + id + ' creado para unidad ' + datos.unidad,
      id,
      '',
      ''
    );

    return { ok: true, id: id };

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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('🎫_Tickets');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de tickets.' };
    }

    var datos = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;
    var estadoActual = '';
    var creadoPor = '';

    for (var i = 1; i < datos.length; i++) {
      if ((datos[i][0] || '').toString().trim() === ticketId) {
        filaReal = i + 1;
        estadoActual = (datos[i][6] || '').toString().trim();
        creadoPor = (datos[i][4] || '').toString().trim();
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
    sheet.getRange(filaReal, 7).setValue('En proceso');    // G: ESTADO
    sheet.getRange(filaReal, 8).setValue(sesion.usuarioId); // H: TECNICO_ASIGNADO
    sheet.getRange(filaReal, 9).setValue(sesion.nombre);   // I: TECNICO_NOMBRE
    sheet.getRange(filaReal, 12).setValue(new Date());      // L: ULTIMA_ACTUALIZACION

    console.log('✅ Ticket tomado:', ticketId, 'por', sesion.nombre);

    // 🔔 NOTIFICACIÓN: Ticket tomado (al creador del ticket)
    if (creadoPor && creadoPor !== sesion.usuarioId) {
      _crearNotificacion(
        creadoPor,
        'TICKET',
        '🔄 El ticket ' + ticketId + ' fue tomado por ' + sesion.nombre,
        ticketId,
        '#panel-tickets'
      );
    }

    // 🔔 NOTIFICACIÓN: Ticket tomado (a Admin/Revisores)
    var admins = _obtenerUsuariosPorRol([1, 2]);
    for (var u = 0; u < admins.length; u++) {
      if (admins[u] !== sesion.usuarioId) {
        _crearNotificacion(
          admins[u],
          'TICKET',
          '🔄 El ticket ' + ticketId + ' fue tomado por ' + sesion.nombre,
          ticketId,
          '#panel-tickets'
        );
      }
    }

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
  // 1. Validar sesión
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }

  const sesion = sesionResp.sesion;

  // 2. Validar permisos (Admin, Revisor o Técnico)
  if (sesion.rol > 3) {
    return { ok: false, error: 'Sin permisos para resolver tickets.' };
  }

  // 3. Validar comentarios
  if (!comentarios || comentarios.length < 5) {
    return { ok: false, error: 'Describe cómo se resolvió (mínimo 5 caracteres).' };
  }

  try {
    // 4. Obtener hoja de Tickets
    const sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Tickets' };
    }

    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];

    // 5. Mapear índices de columnas
    const idxId = headers.indexOf('ID');
    const idxEstado = headers.indexOf('ESTADO');
    const idxFechaCierre = headers.indexOf('FECHA_CIERRE');
    const idxComentarios = headers.indexOf('COMENTARIOS');
    const idxTecnicoAsignado = headers.indexOf('TECNICO_ASIGNADO');
    const idxTecnicoNombre = headers.indexOf('TECNICO_NOMBRE');

    // 6. Verificar columnas necesarias
    if (idxId === -1 || idxEstado === -1) {
      return { ok: false, error: 'Columnas ID o ESTADO no encontradas' };
    }

    // 7. Buscar el ticket
    let filaEncontrada = -1;
    let estadoActual = '';
    let tecnicoAsignado = '';

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxId] === ticketId) {
        filaEncontrada = i + 1;
        estadoActual = datos[i][idxEstado] || '';
        tecnicoAsignado = datos[i][idxTecnicoAsignado] || '';
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { ok: false, error: 'Ticket no encontrado: ' + ticketId };
    }

    // 8. Validar que el ticket esté en proceso (o que sea Admin/Revisor)
    const esAdmin = sesion.rol === 1 || sesion.rol === 2;
    const esTecnicoAsignado = sesion.rol === 3 && tecnicoAsignado === sesion.usuarioId;

    if (!esAdmin && !esTecnicoAsignado) {
      return { ok: false, error: 'No tienes este ticket asignado.' };
    }

    if (!esAdmin && estadoActual !== 'En proceso') {
      return { ok: false, error: 'El ticket debe estar "En proceso" para resolverlo.' };
    }

    // 9. Actualizar estado a "Resuelto"
    sheet.getRange(filaEncontrada, idxEstado + 1).setValue('Resuelto');

    // 10. Registrar fecha de cierre
    if (idxFechaCierre !== -1) {
      sheet.getRange(filaEncontrada, idxFechaCierre + 1).setValue(new Date());
    }

    // 11. Guardar comentarios de resolución
    if (idxComentarios !== -1) {
      const comentariosExistentes = datos[filaEncontrada - 1][idxComentarios] || '';
      const fechaStr = new Date().toLocaleString();
      const nuevoComentario = comentariosExistentes 
        ? comentariosExistentes + '\n--- RESOLUCIÓN ---\n' + comentarios + '\n' + sesion.nombre + ' - ' + fechaStr
        : 'RESOLUCIÓN:\n' + comentarios + '\n' + sesion.nombre + ' - ' + fechaStr;
      sheet.getRange(filaEncontrada, idxComentarios + 1).setValue(nuevoComentario);
    }

    // 12. Registrar en auditoría
    _registrarAuditoria(
      sesion.usuarioId,
      sesion.nombre,
      'RESOLVER_TICKET',
      'TICKETS',
      'Ticket ' + ticketId + ' resuelto por ' + sesion.nombre + '\nComentarios: ' + comentarios,
      ticketId,
      '',
      ''
    );

    // 13. Crear notificación al creador del ticket
    try {
      const idxCreadoPor = headers.indexOf('CREADO_POR');
      if (idxCreadoPor !== -1) {
        const creadorId = datos[filaEncontrada - 1][idxCreadoPor];
        if (creadorId && creadorId !== sesion.usuarioId) {
          _crearNotificacion(
            creadorId,
            'TICKET_RESUELTO',
            '✅ Ticket ' + ticketId + ' ha sido resuelto por ' + sesion.nombre,
            ticketId,
            '#panel-tickets'
          );
        }
      }
    } catch (notifErr) {
      console.warn('⚠️ Error al crear notificación:', notifErr.message);
    }

    console.log('✅ Ticket ' + ticketId + ' resuelto por ' + sesion.nombre);
    
    return { 
      ok: true, 
      mensaje: 'Ticket ' + ticketId + ' resuelto correctamente' 
    };

  } catch (err) {
    console.error('❌ Error al resolver ticket:', err);
    return { ok: false, error: 'Error al resolver ticket: ' + err.message };
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
          var listaEconomicos = economicoAsignado.split(',').map(function (e) { return e.trim(); });
          listaEconomicos.forEach(function (eco) {
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
 * Obtiene los estados de vehículo desde 📋_Catálogo_Estados_Vehiculo
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, estados: [...] }
 */
function obtenerEstadosVehiculo(token) {
  // Validar sesión (opcional, pero seguro)
  if (token) {
    var sesionResp = validarSesion(token);
    if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catálogo_Estados_Vehiculo');
    
    // Si no existe la hoja, crearla con valores por defecto
    if (!sheet) {
      console.log('📋 Hoja de estados no encontrada, creando...');
      sheet = ss.insertSheet('📋_Catálogo_Estados_Vehiculo');
      
      var headers = ['ID', 'NOMBRE', 'COLOR', 'ACTIVO', 'DESCRIPCION'];
      sheet.getRange(1, 1, 1, 5).setValues([headers]);
      
      var datosDefault = [
        [1, 'Activo', 'success', true, 'Vehículo operativo y en circulación'],
        [2, 'Inactivo', 'danger', true, 'Vehículo fuera de circulación temporal o permanente'],
        [3, 'En Mantenimiento', 'warning', true, 'Vehículo en taller por mantenimiento programado'],
        [4, 'Siniestrada', 'dark', true, 'Vehículo con daño por accidente o siniestro']
      ];
      sheet.getRange(2, 1, datosDefault.length, 5).setValues(datosDefault);
      
      sheet.setColumnWidth(1, 50);
      sheet.setColumnWidth(2, 160);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(4, 70);
      sheet.setColumnWidth(5, 280);
      sheet.setFrozenRows(1);
    }

    var datos = sheet.getDataRange().getValues();
    var estados = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue;
      
      // ✅ Verificar que ACTIVO sea true (si existe la columna)
      var activo = true;
      if (fila[3] !== undefined) {
        activo = fila[3] === true || fila[3] === 'TRUE' || fila[3] === 1;
      }
      
      // ✅ Solo incluir estados activos (para el selector)
      if (activo) {
        estados.push({
          id: fila[0].toString(),
          nombre: fila[1] ? fila[1].toString().trim() : '',
          color: fila[2] ? fila[2].toString().trim().toLowerCase() : 'secondary',
          descripcion: fila[4] ? fila[4].toString().trim() : ''
        });
      }
    }

    console.log('📊 Estados cargados:', estados.length);
    return { ok: true, estados: estados };

  } catch (err) {
    console.error('❌ Error en obtenerEstadosVehiculo:', err);
    return { ok: false, error: err.message };
  }
}
// (Admin) Agregar nuevo estado
// 🔧 REEMPLAZAR COMPLETAMENTE - Agregar nuevo estado (Admin/Revisor)
function agregarEstadoVehiculo(token, nombre, color, descripcion) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Admin (1) o Revisor (2)
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para agregar estados.' };
  }

  if (!nombre || nombre.trim() === '') {
    return { ok: false, error: 'El nombre del estado es requerido.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catálogo_Estados_Vehiculo');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados.' };
    }

    var datos = sheet.getDataRange().getValues();

    // Verificar si ya existe un estado con el mismo nombre
    var nombreLimpio = nombre.toString().trim().toUpperCase();
    for (var i = 1; i < datos.length; i++) {
      var existente = (datos[i][1] || '').toString().trim().toUpperCase();
      if (existente === nombreLimpio) {
        return { ok: false, error: 'Ya existe un estado con el nombre "' + nombre + '".' };
      }
    }

    // Calcular nuevo ID
    var ultimoId = 0;
    for (var i = 1; i < datos.length; i++) {
      var id = Number(datos[i][0]) || 0;
      if (id > ultimoId) ultimoId = id;
    }
    var nuevoId = ultimoId + 1;

    // ✅ Guardar con descripción
    sheet.appendRow([
      nuevoId,
      nombre.trim(),
      color || 'secondary',
      true, // ACTIVO
      descripcion || '' // DESCRIPCION
    ]);

    return { ok: true, mensaje: 'Estado "' + nombre + '" agregado correctamente.' };

  } catch (err) {
    console.error('Error en agregarEstadoVehiculo:', err);
    return { ok: false, error: err.message };
  }
}
// (Admin) Editar estado
function editarEstadoVehiculo(token, id, nombre, color, activo, descripcion) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Admin (1) o Revisor (2)
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para editar estados.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del estado es requerido.' };
  }

  if (!nombre || nombre.trim() === '') {
    return { ok: false, error: 'El nombre del estado es requerido.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catálogo_Estados_Vehiculo');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados.' };
    }

    var datos = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;

    for (var i = 1; i < datos.length; i++) {
      if (Number(datos[i][0]) === Number(id)) {
        filaReal = i + 1;
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el estado con ID: ' + id };
    }

    // ✅ Actualizar TODOS los campos incluyendo descripción
    sheet.getRange(filaReal, 2).setValue(nombre.trim());      // NOMBRE
    sheet.getRange(filaReal, 3).setValue(color || 'secondary'); // COLOR
    sheet.getRange(filaReal, 4).setValue(activo === true);     // ACTIVO
    sheet.getRange(filaReal, 5).setValue(descripcion || '');   // DESCRIPCION

    return { ok: true, mensaje: 'Estado actualizado correctamente.' };

  } catch (err) {
    console.error('Error en editarEstadoVehiculo:', err);
    return { ok: false, error: err.message };
  }
}
//Eliminar estado (baja lógica, solo Admin/Revisor)
function eliminarEstadoVehiculo(token, id) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Admin (1) o Revisor (2)
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para eliminar estados.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del estado es requerido.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catálogo_Estados_Vehiculo');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados.' };
    }

    var datos = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;
    var nombreEstado = '';

    for (var i = 1; i < datos.length; i++) {
      if (Number(datos[i][0]) === Number(id)) {
        filaReal = i + 1;
        nombreEstado = datos[i][1] || '';
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el estado con ID: ' + id };
    }

    sheet.getRange(filaReal, 4).setValue(false);

    return { ok: true, mensaje: 'Estado "' + nombreEstado + '" desactivado correctamente.' };

  } catch (err) {
    console.error('Error en eliminarEstadoVehiculo:', err);
    return { ok: false, error: err.message };
  }
}
// ➕ Reactivar estado (Admin/Revisor)
function reactivarEstadoVehiculo(token, id) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Admin (1) o Revisor (2)
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para reactivar estados.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del estado es requerido.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catálogo_Estados_Vehiculo');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados.' };
    }

    var datos = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;
    var nombreEstado = '';

    for (var i = 1; i < datos.length; i++) {
      if (Number(datos[i][0]) === Number(id)) {
        filaReal = i + 1;
        nombreEstado = datos[i][1] || '';
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el estado con ID: ' + id };
    }

    // ✅ REACTIVAR: Cambiar ACTIVO a true
    sheet.getRange(filaReal, 4).setValue(true);

    return { ok: true, mensaje: 'Estado "' + nombreEstado + '" reactivado correctamente.' };

  } catch (err) {
    console.error('Error en reactivarEstadoVehiculo:', err);
    return { ok: false, error: err.message };
  }
}
// ➕ Obtener TODOS los estados (incluyendo inactivos) para el Admin
function obtenerTodosLosEstadosVehiculo(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Admin (1) o Revisor (2) pueden ver todos los estados
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para ver todos los estados.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catálogo_Estados_Vehiculo');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados.' };
    }

    var datos = sheet.getDataRange().getValues();
    var estados = [];

    // ✅ Incluir TODOS los estados (activos e inactivos)
    for (var i = 1; i < datos.length; i++) {
      var id = datos[i][0];
      if (!id) continue;

      estados.push({
        id: Number(id),
        nombre: (datos[i][1] || '').toString().trim(),
        color: (datos[i][2] || 'secondary').toString().trim().toLowerCase(),
        activo: datos[i][3] === true,
        descripcion: (datos[i][4] || '').toString().trim()
      });
    }

    // Ordenar por ID
    estados.sort(function (a, b) { return a.id - b.id; });

    return { ok: true, estados: estados };

  } catch (err) {
    console.error('Error en obtenerTodosLosEstadosVehiculo:', err);
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
// ============================================================
// FUNCIONES PARA GESTIÓN DE ESTADOS DE VEHÍCULOS
// ============================================================

/**
 * Obtiene todos los estados de vehículos disponibles para el selector
 * @returns {Array} - Array de objetos con {id, nombre, color, activo}
 */
function obtenerEstadosVehiculo() {
  try {
    console.log('📋 Obteniendo catálogo de estados de vehículo...');
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('📋_Catálogo_Estados_Vehiculo');

    if (!sheet) {
      console.warn('⚠️ No se encontró la hoja 📋_Catálogo_Estados_Vehiculo');
      // Devolver estados por defecto si no existe la hoja
      return [
        { id: '1', nombre: 'Activo', color: 'success', activo: true },
        { id: '2', nombre: 'Inactivo', color: 'danger', activo: true },
        { id: '3', nombre: 'En Mantenimiento', color: 'warning', activo: true },
        { id: '4', nombre: 'Siniestrada', color: 'dark', activo: true }
      ];
    }

    const data = sheet.getDataRange().getValues();
    const estados = [];

    // Saltar encabezado (fila 1)
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0];          // Columna A: ID
      const nombre = data[i][1];      // Columna B: NOMBRE
      const color = data[i][2];       // Columna C: COLOR
      const activo = data[i][3] === true || data[i][3] === 'TRUE' || data[i][3] === 1;
      const descripcion = data[i][4] || ''; // Columna E: DESCRIPCION

      if (nombre && nombre.toString().trim() !== '') {
        estados.push({
          id: id.toString(),
          nombre: nombre.toString().trim(),
          color: color ? color.toString().trim().toLowerCase() : 'secondary',
          activo: activo,
          descripcion: descripcion
        });
      }
    }

    console.log(`✅ ${estados.length} estados cargados correctamente`);
    return estados;

  } catch (error) {
    console.error('❌ Error al obtener estados de vehículo:', error);
    // Devolver estados por defecto en caso de error
    return [
      { id: '1', nombre: 'Activo', color: 'success', activo: true },
      { id: '2', nombre: 'Inactivo', color: 'danger', activo: true },
      { id: '3', nombre: 'En Mantenimiento', color: 'warning', activo: true },
      { id: '4', nombre: 'Siniestrada', color: 'dark', activo: true }
    ];
  }
}
/**
 * Obtiene el nombre de un estado por su ID
 * @param {string} estadoId - ID del estado
 * @returns {string} - Nombre del estado o 'Sin estado' si no existe
 */
function obtenerNombreEstadoPorId(estadoId) {
  try {
    const estados = obtenerEstadosVehiculo();
    const estado = estados.find(e => e.id === estadoId.toString());
    return estado ? estado.nombre : 'Sin estado';
  } catch (error) {
    console.error('❌ Error al obtener nombre del estado:', error);
    return 'Sin estado';
  }
}
/**
 * Obtiene el color CSS para un estado
 * @param {string} estadoId - ID del estado
 * @returns {string} - Clase de color Bootstrap (success, danger, warning, etc.)
 */
function obtenerColorEstado(estadoId) {
  try {
    const estados = obtenerEstadosVehiculo();
    const estado = estados.find(e => e.id === estadoId.toString());
    return estado ? estado.color : 'secondary';
  } catch (error) {
    console.error('❌ Error al obtener color del estado:', error);
    return 'secondary';
  }
}

/**
 * Guarda o actualiza un vehículo en la flotilla
 */
function guardarVehiculoFlotilla(datos) {
  try {
    console.log('🚚 Guardando vehículo en flotilla:', datos);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Catalogo_Vehiculos');
    if (!sheet) {
      throw new Error('No se encontró la hoja 📋_Catalogo_Vehiculos');
    }
    
    // Validar campos obligatorios
    if (!datos.economico || datos.economico.trim() === '') {
      throw new Error('El económico es obligatorio');
    }
    if (!datos.placas || datos.placas.trim() === '') {
      throw new Error('Las placas son obligatorias');
    }
    if (!datos.estado) {
      throw new Error('El estado es obligatorio');
    }
    if (!datos.empresa) {
      throw new Error('La empresa es obligatoria');
    }
    
    var ahora = new Date();
    var headers = sheet.getDataRange().getValues()[0];
    
    // Obtener índices de columnas
    var idxEconomico = headers.indexOf('ECONOMICO');
    var idxPlacas = headers.indexOf('PLACAS');
    var idxTipoVehiculo = headers.indexOf('TIPO_VEHICULO');
    var idxMarca = headers.indexOf('MARCA');
    var idxModelo = headers.indexOf('MODELO');
    var idxAnio = headers.indexOf('AÑO');
    var idxSerieVehiculo = headers.indexOf('SERIE_VEHICULO');
    var idxGpsActual = headers.indexOf('GPS_ACTUAL');
    var idxEstado = headers.indexOf('ESTADO');
    var idxUltimoServicio = headers.indexOf('ULTIMO_SERVICIO');
    var idxTipoUnidad = headers.indexOf('TIPO_UNIDAD');
    var idxEmpresa = headers.indexOf('EMPRESA');

    // Si no existe la columna EMPRESA, crearla
    if (idxEmpresa === -1) {
      var lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('EMPRESA');
      idxEmpresa = lastCol - 1;
      console.log('✅ Columna EMPRESA creada');
    }
    
    // Buscar si ya existe el vehículo
    var data = sheet.getDataRange().getValues();
    var filaExistente = -1;
    
    for (var i = 1; i < data.length; i++) {
      var econActual = data[i][idxEconomico] ? data[i][idxEconomico].toString().trim() : '';
      if (econActual === datos.economico.trim()) {
        filaExistente = i + 1;
        break;
      }
    }
    
    // Preparar la fila completa
    var nuevaFila = [
      datos.economico.trim(),                    // ECONOMICO
      datos.placas.trim(),                       // PLACAS
      datos.tipoVehiculo || '',                  // TIPO_VEHICULO
      datos.marca || '',                         // MARCA
      datos.modelo || '',                        // MODELO
      datos.anio || '',                          // AÑO
      datos.serieVehiculo || '',                 // SERIE_VEHICULO
      '',                                        // GPS_ACTUAL
      datos.estado || 'Activo',                  // ESTADO
      ahora,                                     // ULTIMO_SERVICIO
      datos.tipoUnidad || '',                    // TIPO_UNIDAD
      datos.empresa || ''                        // ✅ EMPRESA
    ];
    
    if (filaExistente === -1) {
      sheet.appendRow(nuevaFila);
      console.log('✅ Vehículo agregado:', datos.economico);
    } else {
      sheet.getRange(filaExistente, 1, 1, nuevaFila.length).setValues([nuevaFila]);
      console.log('✅ Vehículo actualizado:', datos.economico);
    }
    
    return { ok: true, mensaje: 'Vehículo guardado correctamente' };
    
  } catch (error) {
    console.error('❌ Error al guardar vehículo:', error);
    return { ok: false, error: error.message };
  }
}

// ============================================================
// ADMINISTRACIÓN DE USUARIOS
// ============================================================

/**
 * Obtiene todos los usuarios (incluyendo inactivos)
 * Solo Admin (rol 1) y Revisor (rol 2)
 */
function obtenerUsuarios(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  // Solo Admin o Revisor
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para ver usuarios.' };
  }

  try {
    var sheet = SHEETS.USUARIOS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de usuarios.' };
    }

    var datos = sheet.getDataRange().getValues();
    var usuarios = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue; // Saltar filas vacías

      var rol = Number(fila[4]);
      var rolesLabel = { 1: 'Administrador', 2: 'Revisor', 3: 'Técnico' };

      usuarios.push({
        id: fila[0].toString().trim(),
        nombre: fila[1] ? fila[1].toString().trim() : '',
        usuario: fila[2] ? fila[2].toString().trim() : '',
        rol: rol,
        rolLabel: rolesLabel[rol] || 'Desconocido',
        activo: fila[5] === true,
        ultimoAcceso: fila[6] ? Utilities.formatDate(new Date(fila[6]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '—'
      });
    }

    // Ordenar por ID
    usuarios.sort(function(a, b) { return a.id.localeCompare(b.id); });

    return { ok: true, usuarios: usuarios };

  } catch (err) {
    console.error('Error en obtenerUsuarios:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Agrega un nuevo usuario (solo Admin)
 * @param {string} token - Token de sesión
 * @param {Object} datos - Datos del usuario { nombre, usuario, password, rol, activo }
 * @returns {Object} { ok: true, mensaje: string, id: string }
 */
function agregarUsuario(token, datos) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  // Solo Admin (rol 1) puede agregar usuarios
  if (sesion.rol !== 1) {
    return { ok: false, error: 'Solo Administradores pueden agregar usuarios.' };
  }

  // Validar campos obligatorios
  if (!datos.nombre || datos.nombre.toString().trim() === '') {
    return { ok: false, error: 'El nombre completo es requerido.' };
  }

  if (!datos.usuario || datos.usuario.toString().trim() === '') {
    return { ok: false, error: 'El nombre de usuario es requerido.' };
  }

  if (!datos.password || datos.password.toString().trim().length < 6) {
    return { ok: false, error: 'La contraseña debe tener al menos 6 caracteres.' };
  }

  if (!datos.rol || ![1, 2, 3].includes(Number(datos.rol))) {
    return { ok: false, error: 'Selecciona un rol válido (1=Admin, 2=Revisor, 3=Técnico).' };
  }

  try {
    var sheet = SHEETS.USUARIOS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de usuarios.' };
    }

    var datosSheet = sheet.getDataRange().getValues();

    // Verificar que el usuario no exista (case insensitive)
    var usuarioLimpio = datos.usuario.toString().trim().toLowerCase();
    for (var i = 1; i < datosSheet.length; i++) {
      var existente = (datosSheet[i][2] || '').toString().trim().toLowerCase();
      if (existente === usuarioLimpio) {
        return { ok: false, error: 'El usuario "' + datos.usuario + '" ya existe.' };
      }
    }

    // Generar ID automático según el rol
    var rolesPrefijo = { 1: 'ADM', 2: 'REV', 3: 'TEC' };
    var prefijo = rolesPrefijo[Number(datos.rol)] || 'USR';
    var ultimoId = 0;

    for (var i = 1; i < datosSheet.length; i++) {
      var id = datosSheet[i][0];
      if (id && id.toString().startsWith(prefijo)) {
        var num = parseInt(id.toString().replace(prefijo + '-', '')) || 0;
        if (num > ultimoId) ultimoId = num;
      }
    }
    var nuevoId = prefijo + '-' + String(ultimoId + 1).padStart(3, '0');

    // Hash de la contraseña
    var passHash = hashSimple(datos.password.toString().trim());

    // Activo por defecto (true si no se especifica)
    var activo = (datos.activo !== undefined) ? datos.activo : true;

    // Agregar usuario a la hoja
    sheet.appendRow([
      nuevoId,                                    // A: ID_USUARIO
      datos.nombre.toString().trim(),             // B: NOMBRE
      datos.usuario.toString().trim().toLowerCase(), // C: USUARIO
      passHash,                                   // D: PASS_HASH
      Number(datos.rol),                          // E: ROL
      activo,                                     // F: ACTIVO
      null                                        // G: ULTIMO_ACCESO
    ]);

    console.log('✅ Usuario agregado:', nuevoId, '-', datos.nombre, '(rol:', datos.rol + ')');

    // 🔔 NOTIFICACIÓN: Nuevo usuario creado
    var admins = _obtenerUsuariosPorRol([1]);
    for (var u = 0; u < admins.length; u++) {
      if (admins[u] !== sesion.usuarioId) {
        _crearNotificacion(
          admins[u],
          'USUARIO',
          '👤 Nuevo usuario ' + nuevoId + ' (' + datos.nombre + ') creado con rol ' + datos.rol,
          nuevoId,
          '#panel-usuarios'
        );
      }
    }

    // ✅ AUDITORÍA: Usuario creado
    _registrarAuditoria(
      sesion.usuarioId,
      sesion.nombre,
      'CREAR_USUARIO',
      'USUARIOS',
      'Usuario ' + nuevoId + ' (' + datos.nombre + ') creado con rol ' + datos.rol + ' por ' + sesion.nombre,
      nuevoId,
      '',
      ''
    );

    return {
      ok: true,
      mensaje: 'Usuario "' + datos.nombre + '" agregado correctamente.',
      id: nuevoId
    };

  } catch (err) {
    console.error('❌ Error en agregarUsuario:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Edita un usuario existente (Admin o Revisor)
 * @param {string} token - Token de sesión
 * @param {string} id - ID del usuario a editar
 * @param {Object} datos - Datos a actualizar { nombre, usuario, rol, activo }
 * @returns {Object} { ok: true, mensaje: string }
 */
function editarUsuario(token, id, datos) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  var usuarioActual = sesionResp.sesion;
  
  // Admin (1) o Revisor (2) pueden editar
  if (usuarioActual.rol > 2) {
    return { ok: false, error: 'Sin permisos para editar usuarios.' };
  }

  // Si es Revisor, no puede cambiar el rol
  if (usuarioActual.rol === 2 && datos.rol !== undefined) {
    return { ok: false, error: 'Los Revisores no pueden cambiar el rol de los usuarios.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del usuario es requerido.' };
  }

  try {
    var sheet = SHEETS.USUARIOS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de usuarios.' };
    }

    var datosSheet = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;

    for (var i = 1; i < datosSheet.length; i++) {
      if ((datosSheet[i][0] || '').toString().trim() === id.toString().trim()) {
        filaReal = i + 1;
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el usuario con ID: ' + id };
    }

    // No permitir que un Revisor edite a otro Revisor o Admin
    if (usuarioActual.rol === 2) {
      var rolActual = Number(datosSheet[filaReal - 1][4] || 0);
      if (rolActual <= 2) {
        return { ok: false, error: 'Los Revisores no pueden editar Administradores o Revisores.' };
      }
    }

    // No permitir que un Revisor se edite a sí mismo
    if (usuarioActual.rol === 2 && usuarioActual.usuarioId === id) {
      return { ok: false, error: 'No puedes editar tu propio usuario.' };
    }

    // Actualizar campos
    if (datos.nombre !== undefined) {
      sheet.getRange(filaReal, 2).setValue(datos.nombre.toString().trim());
    }
    
    if (datos.usuario !== undefined) {
      sheet.getRange(filaReal, 3).setValue(datos.usuario.toString().trim().toLowerCase());
    }
    
    // Solo Admin puede cambiar el rol
    if (datos.rol !== undefined && usuarioActual.rol === 1) {
      sheet.getRange(filaReal, 5).setValue(Number(datos.rol));
    }
    
    if (datos.activo !== undefined) {
      sheet.getRange(filaReal, 6).setValue(datos.activo === true);
    }

    console.log('✅ Usuario editado:', id, 'por', usuarioActual.nombre);
    
    return {
      ok: true,
      mensaje: 'Usuario actualizado correctamente.'
    };

  } catch (err) {
    console.error('❌ Error en editarUsuario:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Cambia la contraseña de un usuario (Admin o Revisor)
 * @param {string} token - Token de sesión
 * @param {string} id - ID del usuario
 * @param {string} nuevaPassword - Nueva contraseña (mínimo 6 caracteres)
 * @returns {Object} { ok: true, mensaje: string }
 */
function cambiarPasswordUsuario(token, id, nuevaPassword) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  var usuarioActual = sesionResp.sesion;
  
  // Admin (1) o Revisor (2) pueden cambiar contraseñas
  if (usuarioActual.rol > 2) {
    return { ok: false, error: 'Sin permisos para cambiar contraseñas.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del usuario es requerido.' };
  }

  if (!nuevaPassword || nuevaPassword.toString().trim().length < 6) {
    return { ok: false, error: 'La contraseña debe tener al menos 6 caracteres.' };
  }

  try {
    var sheet = SHEETS.USUARIOS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de usuarios.' };
    }

    var datosSheet = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;

    for (var i = 1; i < datosSheet.length; i++) {
      if ((datosSheet[i][0] || '').toString().trim() === id.toString().trim()) {
        filaReal = i + 1;
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el usuario con ID: ' + id };
    }

    // Revisor no puede cambiar contraseña de Admin o Revisor
    if (usuarioActual.rol === 2) {
      var rolActual = Number(datosSheet[filaReal - 1][4] || 0);
      if (rolActual <= 2) {
        return { ok: false, error: 'Los Revisores no pueden cambiar contraseñas de Administradores o Revisores.' };
      }
    }

    // Hash de la nueva contraseña
    var passHash = hashSimple(nuevaPassword.toString().trim());
    sheet.getRange(filaReal, 4).setValue(passHash);

    console.log('✅ Contraseña cambiada para:', id, 'por', usuarioActual.nombre);
    
    return {
      ok: true,
      mensaje: 'Contraseña actualizada correctamente.'
    };

  } catch (err) {
    console.error('❌ Error en cambiarPasswordUsuario:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Elimina un usuario (baja lógica) - solo Admin
 * @param {string} token - Token de sesión
 * @param {string} id - ID del usuario a eliminar
 * @returns {Object} { ok: true, mensaje: string }
 */
function eliminarUsuario(token, id) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  var usuarioActual = sesionResp.sesion;
  
  // Solo Admin (rol 1) puede eliminar usuarios
  if (usuarioActual.rol !== 1) {
    return { ok: false, error: 'Solo Administradores pueden eliminar usuarios.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del usuario es requerido.' };
  }

  // No permitir eliminar al propio usuario
  if (usuarioActual.usuarioId === id) {
    return { ok: false, error: 'No puedes eliminar tu propio usuario.' };
  }

  try {
    var sheet = SHEETS.USUARIOS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de usuarios.' };
    }

    var datosSheet = sheet.getDataRange().getValues();
    var encontrado = false;
    var filaReal = -1;
    var nombreUsuario = '';

    for (var i = 1; i < datosSheet.length; i++) {
      if ((datosSheet[i][0] || '').toString().trim() === id.toString().trim()) {
        filaReal = i + 1;
        nombreUsuario = datosSheet[i][1] || '';
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { ok: false, error: 'No se encontró el usuario con ID: ' + id };
    }

    // Baja lógica: ACTIVO = false
    sheet.getRange(filaReal, 6).setValue(false);

    console.log('✅ Usuario desactivado:', id, 'por', usuarioActual.nombre);
    
    return {
      ok: true,
      mensaje: 'Usuario "' + nombreUsuario + '" desactivado correctamente.'
    };

  } catch (err) {
    console.error('❌ Error en eliminarUsuario:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// CONFIGURACIÓN DEL SISTEMA
// ============================================================

/**
 * Obtiene todos los parámetros del sistema
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, parametros: [...] }
 */
function obtenerParametros(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  // Solo Admin (1) o Revisor (2)
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para ver configuración.' };
  }

  try {
    var params = _leerParams();
    var parametros = [];

    // Convertir el objeto a array para el frontend
    for (var clave in params) {
      parametros.push({
        clave: clave,
        valor: params[clave] || '',
        descripcion: _obtenerDescripcionParametro(clave)
      });
    }

    // Ordenar alfabéticamente
    parametros.sort(function(a, b) {
      return a.clave.localeCompare(b.clave);
    });

    return { ok: true, parametros: parametros };

  } catch (err) {
    console.error('Error en obtenerParametros:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Obtiene la descripción de un parámetro
 * @param {string} clave - Clave del parámetro
 * @returns {string} Descripción
 */
function _obtenerDescripcionParametro(clave) {
  var descripciones = {
    'FOLIO_PREFIJO': 'Prefijo para folios de servicio',
    'FOLIO_ULTIMO': 'Último folio generado (auto-incrementable)',
    'DRIVE_FOLDER_ID': 'ID de la carpeta raíz en Google Drive',
    'CORREOS_DESTINO': 'Correos para envío de PDF (separados por comas)',
    'DIAS_LIMITE_DRIVE': 'Días de antigüedad para eliminar archivos',
    'COMPRESION_IMAGENES': 'Comprimir imágenes antes de subir (true/false)',
    'CALIDAD_IMAGEN': 'Calidad de compresión (1-100)',
    'MAX_IMAGENES_PDF': 'Máximo de imágenes por PDF',
    'FORMATO_SERIE_GPS': 'Máscara de validación para series GPS',
    'VERBOSE_LOGGING': 'Registro detallado en consola (true/false)',
    'EMPRESA_NOMBRE': 'Nombre de la empresa para documentos',
    'EMPRESA_RFC': 'RFC para facturación',
    'TICKET_PREFIJO': 'Prefijo para tickets de fallas',
    'TICKET_ULTIMO': 'Último ticket generado'
  };
  return descripciones[clave] || '';
}

/**
 * Actualiza un parámetro del sistema
 * @param {string} token - Token de sesión
 * @param {string} clave - Clave del parámetro
 * @param {string} valor - Nuevo valor
 * @returns {Object} { ok: true, mensaje: string }
 */
function actualizarParametro(token, clave, valor) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  // Solo Admin (1) o Revisor (2)
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para modificar configuración.' };
  }

  if (!clave) {
    return { ok: false, error: 'La clave del parámetro es requerida.' };
  }

  try {
    // Validar valores según el tipo
    if (clave === 'COMPRESION_IMAGENES') {
      var valorBool = valor.toString().toLowerCase().trim();
      if (valorBool !== 'true' && valorBool !== 'false') {
        return { ok: false, error: 'COMPRESION_IMAGENES debe ser true o false.' };
      }
    }

    if (clave === 'CALIDAD_IMAGEN') {
      var calidad = Number(valor);
      if (isNaN(calidad) || calidad < 1 || calidad > 100) {
        return { ok: false, error: 'CALIDAD_IMAGEN debe ser un número entre 1 y 100.' };
      }
    }

    if (clave === 'CORREOS_DESTINO') {
      var correos = valor.toString().split(',').map(function(c) { return c.trim(); });
      var validos = correos.filter(function(c) { return c.includes('@'); });
      if (validos.length === 0 && valor.toString().trim() !== '') {
        return { ok: false, error: 'Formato de correos inválido. Usa: correo1@dominio.com, correo2@dominio.com' };
      }
    }

    // Guardar el parámetro
    _escribirParam(clave, valor.toString().trim());

    console.log('✅ Parámetro actualizado:', clave, '=', valor);

    return {
      ok: true,
      mensaje: 'Parámetro "' + clave + '" actualizado correctamente.'
    };

  } catch (err) {
    console.error('Error en actualizarParametro:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Actualiza múltiples parámetros a la vez
 * @param {string} token - Token de sesión
 * @param {Array} parametros - Array de { clave, valor }
 * @returns {Object} { ok: true, mensaje: string }
 */
function actualizarMultiplesParametros(token, parametros) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para modificar configuración.' };
  }

  if (!parametros || parametros.length === 0) {
    return { ok: false, error: 'No hay parámetros para actualizar.' };
  }

  try {
    var actualizados = 0;
    var errores = [];

    for (var i = 0; i < parametros.length; i++) {
      var p = parametros[i];
      try {
        _escribirParam(p.clave, p.valor.toString().trim());
        actualizados++;
      } catch (err) {
        errores.push(p.clave + ': ' + err.message);
      }
    }

    if (errores.length > 0) {
      return {
        ok: true,
        mensaje: 'Se actualizaron ' + actualizados + ' parámetros. Errores: ' + errores.join(', ')
      };
    }

    return {
      ok: true,
      mensaje: actualizados + ' parámetros actualizados correctamente.'
    };

  } catch (err) {
    console.error('Error en actualizarMultiplesParametros:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// SISTEMA DE NOTIFICACIONES
// ============================================================

/**
 * Crea una nueva notificación en el sistema
 * @param {string} usuarioDestino - ID del usuario destinatario
 * @param {string} tipo - Tipo de notificación (REPORTE, TICKET, USUARIO, etc.)
 * @param {string} mensaje - Mensaje descriptivo
 * @param {string} folioRelacionado - Folio o ID relacionado (opcional)
 * @param {string} urlAccion - URL para acción rápida (opcional)
 */
function _crearNotificacion(usuarioId, tipo, mensaje, folioRelacionado, urlAccion) {
  try {
    const sheet = SHEETS.NOTIFICACIONES();
    if (!sheet) {
      console.warn('⚠️ No se encontró la hoja de Notificaciones');
      return;
    }

    const nuevaFila = sheet.getLastRow() + 1;
    const datos = [
      new Date(),           // FECHA
      usuarioId,            // USUARIO_DESTINO
      tipo,                 // TIPO
      mensaje,              // MENSAJE
      folioRelacionado || '', // FOLIO_RELACIONADO
      false,                // LEIDA
      null,                 // FECHA_LECTURA
      urlAccion || ''       // URL_ACCION
    ];

    sheet.getRange(nuevaFila, 1, 1, datos.length).setValues([datos]);
    console.log('✅ Notificación creada para:', usuarioId);

  } catch (err) {
    console.error('❌ Error al crear notificación:', err);
  }
}

/**
 * Obtiene las notificaciones del usuario actual
 * @param {string} token - Token de sesión
 * @param {boolean} soloNoLeidas - Si solo obtener no leídas
 * @returns {Object} { ok: true, notificaciones: [...] }
 */
function obtenerNotificaciones(token, soloNoLeidas) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  var usuarioId = sesionResp.sesion.usuarioId;

  try {
    var sheet = SHEETS.NOTIFICACIONES();
    if (!sheet) {
      return { ok: true, notificaciones: [] };
    }

    var datos = sheet.getDataRange().getValues();
    var notificaciones = [];

    var tiposIcono = {
      'REPORTE': '📝',
      'TICKET': '🎫',
      'USUARIO': '👤',
      'ESTADO': '🔄',
      'APROBACION': '✅',
      'PAGO': '💰',
      'SISTEMA': '⚙️'
    };

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var destino = (fila[1] || '').toString().trim();

      if (destino !== usuarioId) continue;

      var leida = fila[5] === true;
      if (soloNoLeidas && leida) continue;

      var tipo = (fila[2] || 'SISTEMA').toString().trim();
      var icono = tiposIcono[tipo] || '📌';

      notificaciones.push({
        id: i, // Usamos la fila como ID
        fecha: fila[0] ? Utilities.formatDate(new Date(fila[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '—',
        tipo: tipo,
        mensaje: fila[3] ? fila[3].toString().trim() : '',
        folioRelacionado: fila[4] ? fila[4].toString().trim() : '',
        leida: leida,
        fechaLectura: fila[6] ? Utilities.formatDate(new Date(fila[6]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
        urlAccion: fila[7] ? fila[7].toString().trim() : '',
        icono: icono
      });
    }

    // Ordenar por fecha (más reciente primero)
    notificaciones.sort(function(a, b) {
      return new Date(b.fecha) - new Date(a.fecha);
    });

    return {
      ok: true,
      notificaciones: notificaciones,
      noLeidas: notificaciones.filter(function(n) { return !n.leida; }).length
    };

  } catch (err) {
    console.error('Error en obtenerNotificaciones:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Marca una notificación como leída
 * @param {string} token - Token de sesión
 * @param {number} notificacionId - ID de la notificación (fila)
 * @returns {Object} { ok: true }
 */
function marcarNotificacionLeida(token, notificacionId) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var sheet = SHEETS.NOTIFICACIONES();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de notificaciones.' };
    }

    // La fila real es notificacionId + 1 (porque el ID es el índice del array)
    var filaReal = Number(notificacionId) + 1;

    // Verificar que la fila existe
    var ultimaFila = sheet.getLastRow();
    if (filaReal > ultimaFila) {
      return { ok: false, error: 'Notificación no encontrada.' };
    }

    // Verificar que pertenece al usuario
    var destino = sheet.getRange(filaReal, 2).getValue();
    if (destino.toString().trim() !== sesionResp.sesion.usuarioId) {
      return { ok: false, error: 'No tienes permiso para modificar esta notificación.' };
    }

    // Marcar como leída
    sheet.getRange(filaReal, 6).setValue(true);  // LEIDA
    sheet.getRange(filaReal, 7).setValue(new Date()); // FECHA_LECTURA

    return { ok: true };

  } catch (err) {
    console.error('Error en marcarNotificacionLeida:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Marca todas las notificaciones del usuario como leídas
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true }
 */
function marcarTodasNotificacionesLeidas(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  var usuarioId = sesionResp.sesion.usuarioId;

  try {
    var sheet = SHEETS.NOTIFICACIONES();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de notificaciones.' };
    }

    var datos = sheet.getDataRange().getValues();
    var ahora = new Date();

    for (var i = 1; i < datos.length; i++) {
      var destino = (datos[i][1] || '').toString().trim();
      var leida = datos[i][5] === true;

      if (destino === usuarioId && !leida) {
        sheet.getRange(i + 1, 6).setValue(true);
        sheet.getRange(i + 1, 7).setValue(ahora);
      }
    }

    return { ok: true };

  } catch (err) {
    console.error('Error en marcarTodasNotificacionesLeidas:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Obtiene los IDs de usuarios con un rol específico
 * @param {Array} roles - Array de roles [1, 2, 3]
 * @returns {Array} - Array de IDs de usuarios
 */
function _obtenerUsuariosPorRol(roles) {
  try {
    var sheet = SHEETS.USUARIOS();
    if (!sheet) return [];

    var datos = sheet.getDataRange().getValues();
    var usuarios = [];

    for (var i = 1; i < datos.length; i++) {
      var rol = Number(datos[i][4] || 0);
      var activo = datos[i][5] === true;
      var id = (datos[i][0] || '').toString().trim();

      if (id && activo && roles.indexOf(rol) !== -1) {
        usuarios.push(id);
      }
    }

    return usuarios;

  } catch (err) {
    console.error('Error en _obtenerUsuariosPorRol:', err);
    return [];
  }
}
/**
 * FUNCIÓN DE PRUEBA - Crear una notificación manual para verificar que funciona
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, mensaje: string }
 */
function pruebaCrearNotificacion(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  var usuarioId = sesionResp.sesion.usuarioId;
  var nombre = sesionResp.sesion.nombre;
  
  // Crear una notificación de prueba
  _crearNotificacion(
    usuarioId,
    'SISTEMA',
    '🧪 Notificación de prueba para ' + nombre + ' - El sistema funciona correctamente',
    'PRUEBA-' + new Date().getTime(),
    '#panel-dashboard'
  );
  
  return { ok: true, mensaje: 'Notificación de prueba creada para ' + nombre };
}
// ============================================================
// LOG DE AUDITORÍA
// ============================================================

/**
 * Registra una acción en el log de auditoría
 * @param {string} usuarioId - ID del usuario que realiza la acción
 * @param {string} usuarioNombre - Nombre del usuario
 * @param {string} accion - Tipo de acción (LOGIN, REPORTE, TICKET, etc.)
 * @param {string} modulo - Módulo afectado (USUARIOS, INVENTARIO, FLOTILLA, etc.)
 * @param {string} descripcion - Descripción detallada de la acción
 * @param {string} folioRelacionado - Folio o ID relacionado (opcional)
 * @param {string} ip - Dirección IP (opcional)
 * @param {string} userAgent - Navegador/Dispositivo (opcional)
 */
function _registrarAuditoria(usuarioId, usuarioNombre, accion, modulo, descripcion, folioRelacionado, ip, userAgent) {
  try {
    var sheet = SS.getSheetByName('📈_Log_Auditoria');
    if (!sheet) {
      console.warn('⚠️ No se encontró la hoja de auditoría');
      return;
    }

    var ahora = new Date();

    // ✅ ORDEN CORRECTO: 10 columnas
    sheet.appendRow([
      ahora,                    // 1. FECHA
      usuarioId,                // 2. USUARIO_ID
      usuarioNombre,            // 3. USUARIO_NOMBRE
      accion,                   // 4. ACCION
      modulo,                   // 5. MODULO
      descripcion,              // 6. DESCRIPCION
      folioRelacionado || '',   // 7. FOLIO_RELACIONADO
      ip || '',                 // 8. IP
      userAgent || '',          // 9. USER_AGENT
      ''                        // 10. DETALLES_ADICIONALES
    ]);

    console.log('📝 Auditoría registrada:', accion, '-', usuarioNombre);

  } catch (err) {
    console.error('❌ Error al registrar auditoría:', err);
  }
}

/**
 * Obtiene el log de auditoría con filtros
 * @param {string} token - Token de sesión
 * @param {Object} filtros - { usuario, accion, modulo, fechaDesde, fechaHasta, limite }
 * @returns {Object} { ok: true, registros: [...] }
 */
function obtenerLogAuditoria(token, filtros) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  // Solo Admin y Revisor pueden ver el log
  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para ver el log de auditoría.' };
  }

  try {
    var sheet = SS.getSheetByName('📈_Log_Auditoria');
    if (!sheet) {
      return { ok: true, registros: [] };
    }

    var datos = sheet.getDataRange().getValues();
    if (datos.length <= 1) {
      return { ok: true, registros: [] };
    }

    var headers = datos[0];
    var idxFecha = headers.indexOf('FECHA');
    var idxUsuarioId = headers.indexOf('USUARIO_ID');
    var idxUsuarioNombre = headers.indexOf('USUARIO_NOMBRE');
    var idxAccion = headers.indexOf('ACCION');
    var idxModulo = headers.indexOf('MODULO');
    var idxDescripcion = headers.indexOf('DESCRIPCION');
    var idxFolio = headers.indexOf('FOLIO_RELACIONADO');
    var idxIP = headers.indexOf('IP');
    var idxUserAgent = headers.indexOf('USER_AGENT');

    var limite = filtros.limite || 100;
    var registros = [];

    for (var i = datos.length - 1; i >= 1; i--) {
      if (registros.length >= limite) break;

      var fila = datos[i];
      if (!fila[idxFecha]) continue;

      // Aplicar filtros
      if (filtros.usuario && fila[idxUsuarioId] !== filtros.usuario) continue;
      if (filtros.accion && fila[idxAccion] !== filtros.accion) continue;
      if (filtros.modulo && fila[idxModulo] !== filtros.modulo) continue;

      if (filtros.fechaDesde) {
        var fechaDesde = new Date(filtros.fechaDesde);
        var fechaRegistro = new Date(fila[idxFecha]);
        if (fechaRegistro < fechaDesde) continue;
      }
      if (filtros.fechaHasta) {
        var fechaHasta = new Date(filtros.fechaHasta);
        var fechaRegistro = new Date(fila[idxFecha]);
        if (fechaRegistro > fechaHasta) continue;
      }

      registros.push({
        fecha: fila[idxFecha] ? Utilities.formatDate(new Date(fila[idxFecha]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss') : '—',
        usuarioId: fila[idxUsuarioId] || '',
        usuarioNombre: fila[idxUsuarioNombre] || '',
        accion: fila[idxAccion] || '',
        modulo: fila[idxModulo] || '',
        descripcion: fila[idxDescripcion] || '',
        folioRelacionado: fila[idxFolio] || '',
        ip: fila[idxIP] || '',
        userAgent: fila[idxUserAgent] || ''
      });
    }

    return { ok: true, registros: registros };

  } catch (err) {
    console.error('Error en obtenerLogAuditoria:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Obtiene estadísticas del log de auditoría
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, stats: {...} }
 */
function obtenerEstadisticasAuditoria(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos.' };
  }

  try {
    var sheet = SS.getSheetByName('📈_Log_Auditoria');
    if (!sheet) {
      return { ok: true, stats: { total: 0, porAccion: {}, porModulo: {}, porUsuario: {} } };
    }

    var datos = sheet.getDataRange().getValues();
    var stats = {
      total: datos.length - 1,
      porAccion: {},
      porModulo: {},
      porUsuario: {},
      hoy: 0,
      estaSemana: 0,
      esteMes: 0
    };

    var ahora = new Date();
    var hoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
    var semanaInicio = new Date(ahora);
    semanaInicio.setDate(semanaInicio.getDate() - 7);
    var mesInicio = new Date(ahora.getFullYear(), ahora.getMonth(), 1);

    var headers = datos[0];
    var idxFecha = headers.indexOf('FECHA');
    var idxAccion = headers.indexOf('ACCION');
    var idxModulo = headers.indexOf('MODULO');
    var idxUsuarioNombre = headers.indexOf('USUARIO_NOMBRE');

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[idxFecha]) continue;

      var fecha = new Date(fila[idxFecha]);
      
      // Contar por período
      if (fecha >= hoy) stats.hoy++;
      if (fecha >= semanaInicio) stats.estaSemana++;
      if (fecha >= mesInicio) stats.esteMes++;

      // Contar por acción
      var accion = fila[idxAccion] || 'DESCONOCIDO';
      stats.porAccion[accion] = (stats.porAccion[accion] || 0) + 1;

      // Contar por módulo
      var modulo = fila[idxModulo] || 'DESCONOCIDO';
      stats.porModulo[modulo] = (stats.porModulo[modulo] || 0) + 1;

      // Contar por usuario
      var usuario = fila[idxUsuarioNombre] || 'DESCONOCIDO';
      stats.porUsuario[usuario] = (stats.porUsuario[usuario] || 0) + 1;
    }

    return { ok: true, stats: stats };

  } catch (err) {
    console.error('Error en obtenerEstadisticasAuditoria:', err);
    return { ok: false, error: err.message };
  }
}
// 
function registrarAccesoFrontend(token, userAgent) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  var sesion = sesionResp.sesion;

  _registrarAuditoria(
    sesion.usuarioId,
    sesion.nombre,
    'ACCESO_SISTEMA',
    'AUTENTICACION',
    'Acceso al sistema desde la web app',
    '',
    '',
    userAgent || ''
  );

  return { ok: true };
}
/**
 * Crear registros de auditoría de prueba
 * Ejecutar una vez en el editor de Apps Script
 */
function crearAuditoriaPruebaV2() {
  console.log('📝 Creando registros de auditoría de prueba...');
  
  // Obtener usuarios activos
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('👤_Usuarios');
  if (!sheet) {
    console.error('❌ No se encontró la hoja de usuarios');
    return;
  }
  
  var datos = sheet.getDataRange().getValues();
  var usuarios = [];
  
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][5] === true) {
      usuarios.push({
        id: datos[i][0] || '',
        nombre: datos[i][1] || 'Usuario'
      });
    }
  }
  
  if (usuarios.length === 0) {
    console.log('⚠️ No hay usuarios activos');
    return;
  }
  
  var acciones = [
    { accion: 'LOGIN_EXITOSO', modulo: 'AUTENTICACION' },
    { accion: 'CREAR_REPORTE', modulo: 'REPORTES' },
    { accion: 'APROBAR_REPORTE', modulo: 'REPORTES' },
    { accion: 'CREAR_TICKET', modulo: 'TICKETS' },
    { accion: 'TOMAR_TICKET', modulo: 'TICKETS' },
    { accion: 'RESOLVER_TICKET', modulo: 'TICKETS' },
    { accion: 'CREAR_USUARIO', modulo: 'USUARIOS' }
  ];
  
  var contador = 0;
  
  usuarios.forEach(function(u) {
    acciones.forEach(function(a) {
      contador++;
      var descripcion = a.accion + ' - ' + u.nombre + ' (Prueba ' + contador + ')';
      var folio = 'PRUEBA-' + String(contador).padStart(4, '0');
      
      _registrarAuditoria(
        u.id,
        u.nombre,
        a.accion,
        a.modulo,
        descripcion,
        folio,
        '192.168.1.' + contador,
        'Mozilla/5.0 (Prueba)'
      );
    });
  });
  
  console.log('✅ Creados ' + contador + ' registros de auditoría de prueba');
}
function verEstructuraAuditoria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('📈_Log_Auditoria');
  
  if (!sheet) {
    console.log('❌ No se encontró la hoja');
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  console.log('📋 Encabezados actuales:', headers);
  console.log('📋 Número de columnas:', headers.length);
}
function recrearHojaAuditoria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('📈_Log_Auditoria');
  
  // Si existe, eliminarla y crearla de nuevo
  if (sheet) {
    ss.deleteSheet(sheet);
    console.log('🗑️ Hoja antigua eliminada');
  }
  
  // Crear nueva hoja
  sheet = ss.insertSheet('📈_Log_Auditoria');
  
  // Encabezados correctos (10 columnas)
  var headers = [
    'FECHA',
    'USUARIO_ID',
    'USUARIO_NOMBRE',
    'ACCION',
    'MODULO',
    'DESCRIPCION',
    'FOLIO_RELACIONADO',
    'IP',
    'USER_AGENT',
    'DETALLES_ADICIONALES'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Ajustar anchos
  sheet.setColumnWidth(1, 150); // FECHA
  sheet.setColumnWidth(2, 120); // USUARIO_ID
  sheet.setColumnWidth(3, 150); // USUARIO_NOMBRE
  sheet.setColumnWidth(4, 150); // ACCION
  sheet.setColumnWidth(5, 120); // MODULO
  sheet.setColumnWidth(6, 250); // DESCRIPCION
  sheet.setColumnWidth(7, 130); // FOLIO_RELACIONADO
  sheet.setColumnWidth(8, 120); // IP
  sheet.setColumnWidth(9, 200); // USER_AGENT
  sheet.setColumnWidth(10, 200); // DETALLES_ADICIONALES
  
  // Congelar primera fila
  sheet.setFrozenRows(1);
  
  console.log('✅ Hoja 📈_Log_Auditoria recreada correctamente');
  console.log('📋 Encabezados:', headers.join(' | '));
}
function limpiarAuditoriaPrueba() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('📈_Log_Auditoria');
  if (!sheet) {
    console.log('❌ No se encontró la hoja');
    return;
  }
  
  var datos = sheet.getDataRange().getValues();
  var filasAEliminar = [];
  
  for (var i = datos.length - 1; i >= 1; i--) {
    var fila = datos[i];
    var folio = fila[6] || ''; // FOLIO_RELACIONADO
    if (folio.toString().startsWith('PRUEBA-')) {
      sheet.deleteRow(i + 1);
    }
  }
  
  console.log('✅ Registros de prueba eliminados');
}
/**
 * Exporta la flotilla a Excel (SIN guardar en Drive)
 * @param {string} token - Token de sesión
 * @param {string} filtroEmpresa - 'TEG', 'ALVA' o '' (todas)
 * @returns {Object} { ok: true, base64: string, nombre: string }
 */
function exportarFlotillaExcel(token, filtroEmpresa) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para exportar flotilla.' };
  }

  try {
    var resp = obtenerFlotillaCompleta(token, {});
    if (!resp.ok) return { ok: false, error: resp.error };
    var vehiculos = resp.vehiculos || [];

    if (filtroEmpresa) {
      vehiculos = vehiculos.filter(function(v) { return v.empresa === filtroEmpresa; });
    }

    var empresas = {};
    vehiculos.forEach(function(v) {
      var empresa = v.empresa || 'SIN EMPRESA';
      if (!empresas[empresa]) { empresas[empresa] = []; }
      empresas[empresa].push(v);
    });

    var html = generarHTMLFlotillaExcel(empresas);
    var nombreArchivo = 'Flotilla_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + (filtroEmpresa ? '_' + filtroEmpresa : '') + '.xls';
    
    // ✅ GENERAR BLOB EN MEMORIA (NO SE GUARDA EN DRIVE)
    var blob = Utilities.newBlob(html, 'application/vnd.ms-excel', nombreArchivo);
    var base64Data = Utilities.base64Encode(blob.getBytes());

    console.log('✅ Excel generado en memoria:', nombreArchivo);
    
    return { 
      ok: true, 
      base64: base64Data, 
      nombre: nombreArchivo 
    };

  } catch (err) {
    console.error('Error en exportarFlotillaExcel:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Genera el HTML para el Excel de flotilla
 */
function generarHTMLFlotillaExcel(empresas) {
  var html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" 
      xmlns:x="urn:schemas-microsoft-com:office:excel" 
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8">
  <!--[if gte mso 9]>
  <xml>
    <x:ExcelWorkbook>
      <x:ExcelWorksheets>
        <x:ExcelWorksheet>
          <x:Name>Flotilla</x:Name>
          <x:WorksheetOptions>
            <x:DisplayGridlines/>
          </x:WorksheetOptions>
        </x:ExcelWorksheet>
      </x:ExcelWorksheets>
    </x:ExcelWorkbook>
  </xml>
  <![endif]-->
  <style>
    table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; width: 100%; }
    th { background: #1a56db; color: #ffffff; font-weight: bold; padding: 8px 12px; border: 1px solid #000; text-align: left; }
    td { padding: 6px 12px; border: 1px solid #ccc; }
    .header { font-size: 14pt; font-weight: bold; margin-bottom: 10px; color: #1a56db; }
    .fecha { font-size: 9pt; color: #666; margin-bottom: 15px; }
    .resumen-empresa { background: #f0f4ff; font-weight: bold; }
    .resumen-total { background: #e8f0fe; font-weight: bold; }
    .accesorio-item { padding-left: 20px; }
  </style>
</head>
<body>
  <div class="header">🚚 REPORTE DE FLOTILLA POR EMPRESA</div>
  <div class="fecha">Fecha de exportación: ${new Date().toLocaleString('es-MX')}</div>`;

  var totalGeneral = {
    gateways: 0,
    camaras: 0,
    accesorios: {}
  };

  var empresasOrdenadas = Object.keys(empresas).sort();

  empresasOrdenadas.forEach(function(empresa) {
    var vehiculos = empresas[empresa];
    var resumenEmpresa = {
      gateways: 0,
      camaras: 0,
      accesorios: {}
    };

    html += `
  <br><br>
  <div style="font-size: 12pt; font-weight: bold; background: #0f172a; color: #fff; padding: 8px 12px; border-radius: 4px;">
    🏢 EMPRESA: ${empresa}
  </div>
  <table>
    <thead>
      <tr>
        <th>Económico</th>
        <th>Placas</th>
        <th>Tipo</th>
        <th>Marca/Modelo</th>
        <th>Estado</th>
        <th>Tipo Unidad</th>
        <th>📡 Gateway</th>
        <th>📷 Cámara</th>
        <th>🔧 Accesorios</th>
      </tr>
    </thead>
    <tbody>`;

    vehiculos.forEach(function(v) {
      var accesoriosStr = (v.accesorios && v.accesorios.length > 0) 
        ? v.accesorios.join(', ') 
        : '—';

      // Contar equipos
      if (v.gatewaySerie) resumenEmpresa.gateways++;
      if (v.camaraSerie) resumenEmpresa.camaras++;

      // Contar accesorios individuales
      if (v.accesorios && v.accesorios.length > 0) {
        v.accesorios.forEach(function(a) {
          resumenEmpresa.accesorios[a] = (resumenEmpresa.accesorios[a] || 0) + 1;
        });
      }

      html += `
      <tr>
        <td><strong>${v.economico || '—'}</strong></td>
        <td>${v.placas || '—'}</td>
        <td>${v.tipoVehiculo || '—'}</td>
        <td>${(v.marca || '')} ${(v.modelo || '')}</td>
        <td>${v.estado || 'Sin estado'}</td>
        <td>${v.tipoUnidad || '—'}</td>
        <td>${v.gatewaySerie || '—'}</td>
        <td>${v.camaraSerie || '—'}</td>
        <td style="font-size: 9pt;">${accesoriosStr}</td>
      </tr>`;
    });

    html += `
    </tbody>
  </table>`;

    // Resumen por empresa
    html += `
  <div style="margin-top: 8px; background: #f0f4ff; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px;">
    <strong>📊 RESUMEN ${empresa}:</strong><br>
    &nbsp;&nbsp;🔵 Gateways: ${resumenEmpresa.gateways}<br>
    &nbsp;&nbsp;📷 Cámaras: ${resumenEmpresa.camaras}<br>
    &nbsp;&nbsp;🔧 Accesorios:`;

    var accesoriosKeys = Object.keys(resumenEmpresa.accesorios);
    if (accesoriosKeys.length === 0) {
      html += ` Ninguno`;
    } else {
      html += `<br>`;
      accesoriosKeys.sort().forEach(function(a) {
        html += `&nbsp;&nbsp;&nbsp;&nbsp;• ${a}: ${resumenEmpresa.accesorios[a]}<br>`;
        // Acumular para total general
        totalGeneral.accesorios[a] = (totalGeneral.accesorios[a] || 0) + resumenEmpresa.accesorios[a];
      });
    }
    html += `</div>`;

    // Acumular totales generales
    totalGeneral.gateways += resumenEmpresa.gateways;
    totalGeneral.camaras += resumenEmpresa.camaras;
  });

  // Total general
  html += `
  <br><br>
  <div style="font-size: 12pt; font-weight: bold; background: #0f172a; color: #fff; padding: 8px 12px; border-radius: 4px;">
    📋 TOTAL GENERAL
  </div>
  <div style="background: #e8f0fe; padding: 8px 12px; border: 1px solid #1a56db; border-radius: 4px;">
    &nbsp;&nbsp;🔵 Gateways: ${totalGeneral.gateways}<br>
    &nbsp;&nbsp;📷 Cámaras: ${totalGeneral.camaras}<br>
    &nbsp;&nbsp;🔧 Accesorios:`;

  var accesoriosKeysTotal = Object.keys(totalGeneral.accesorios);
  if (accesoriosKeysTotal.length === 0) {
    html += ` Ninguno`;
  } else {
    html += `<br>`;
    accesoriosKeysTotal.sort().forEach(function(a) {
      html += `&nbsp;&nbsp;&nbsp;&nbsp;• ${a}: ${totalGeneral.accesorios[a]}<br>`;
    });
  }
  html += `</div>`;

  html += `
  <div style="margin-top: 15px; font-size: 8pt; color: #999;">
    Documento generado automáticamente por Fleet Manager - ${new Date().toLocaleString('es-MX')}
  </div>
</body>
</html>`;

  return html;
}
// ============================================================
// ESTADOS DE INVENTARIO - CATÁLOGO DINÁMICO
// ============================================================

/**
 * Obtiene todos los estados de inventario desde 📋_Estados_Inventario
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, estados: [...] }
 */
function obtenerEstadosInventario(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Estados_Inventario');
    
    // Si no existe la hoja, crearla con valores por defecto
    if (!sheet) {
      sheet = ss.insertSheet('📋_Estados_Inventario');
      var headers = ['ID', 'NOMBRE', 'COLOR'];
      sheet.getRange(1, 1, 1, 3).setValues([headers]);
      
      var datosDefault = [
        [1, 'Disponible', 'success'],
        [2, 'Instalado', 'primary'],
        [3, 'Garantía', 'warning'],
        [4, 'Baja', 'danger']
      ];
      sheet.getRange(2, 1, datosDefault.length, 3).setValues(datosDefault);
      
      // Ajustar anchos
      sheet.setColumnWidth(1, 60);
      sheet.setColumnWidth(2, 120);
      sheet.setColumnWidth(3, 100);
      
      return { ok: true, estados: datosDefault.map(function(e) {
        return { id: e[0], nombre: e[1], color: e[2] };
      })};
    }

    var datos = sheet.getDataRange().getValues();
    var estados = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue;
      
      estados.push({
        id: Number(fila[0]),
        nombre: fila[1] ? fila[1].toString().trim() : '',
        color: fila[2] ? fila[2].toString().trim().toLowerCase() : 'secondary'
      });
    }

    return { ok: true, estados: estados };

  } catch (err) {
    console.error('Error en obtenerEstadosInventario:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Agrega un nuevo estado de inventario (solo Admin)
 * @param {string} token - Token de sesión
 * @param {string} nombre - Nombre del estado
 * @param {string} color - Color (success, danger, warning, etc.)
 * @returns {Object} { ok: true }
 */
function agregarEstadoInventario(token, nombre, color) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (sesionResp.sesion.rol !== 1) {
    return { ok: false, error: 'Solo Administradores pueden agregar estados de inventario.' };
  }

  if (!nombre || nombre.trim() === '') {
    return { ok: false, error: 'El nombre del estado es requerido.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Estados_Inventario');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados de inventario.' };
    }

    var datos = sheet.getDataRange().getValues();
    
    // Calcular nuevo ID
    var ultimoId = 0;
    for (var i = 1; i < datos.length; i++) {
      var id = Number(datos[i][0]) || 0;
      if (id > ultimoId) ultimoId = id;
    }
    var nuevoId = ultimoId + 1;

    sheet.appendRow([
      nuevoId,
      nombre.trim(),
      color || 'secondary'
    ]);

    return { ok: true, mensaje: 'Estado "' + nombre + '" agregado correctamente.' };

  } catch (err) {
    console.error('Error en agregarEstadoInventario:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Elimina un estado de inventario (solo Admin)
 * @param {string} token - Token de sesión
 * @param {number} id - ID del estado a eliminar
 * @returns {Object} { ok: true }
 */
function eliminarEstadoInventario(token, id) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (sesionResp.sesion.rol !== 1) {
    return { ok: false, error: 'Solo Administradores pueden eliminar estados de inventario.' };
  }

  if (!id) {
    return { ok: false, error: 'El ID del estado es requerido.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('📋_Estados_Inventario');
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de estados de inventario.' };
    }

    var datos = sheet.getDataRange().getValues();
    var filaEliminar = -1;

    for (var i = 1; i < datos.length; i++) {
      if (Number(datos[i][0]) === Number(id)) {
        filaEliminar = i + 1;
        break;
      }
    }

    if (filaEliminar === -1) {
      return { ok: false, error: 'No se encontró el estado con ID: ' + id };
    }

    sheet.deleteRow(filaEliminar);

    return { ok: true, mensaje: 'Estado eliminado correctamente.' };

  } catch (err) {
    console.error('Error en eliminarEstadoInventario:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Obtiene alertas del sistema para el dashboard
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, alertas: [...] }
 */
function obtenerAlertas(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };
  
  var esAdminRevisor = sesionResp.sesion.rol <= 2;
  var alertas = [];
  
  try {
    // 1. Vehículos sin GPS (solo Admin/Revisor)
    if (esAdminRevisor) {
      var flotillaResp = obtenerFlotillaCompleta(token, {});
      if (flotillaResp.ok && flotillaResp.vehiculos) {
        var sinGPS = flotillaResp.vehiculos.filter(function(v) {
          return v.estado === 'Activo' && !v.gatewaySerie;
        });
        if (sinGPS.length > 0) {
          alertas.push({
            tipo: 'warning',
            icono: '📡',
            mensaje: sinGPS.length + ' vehículo' + (sinGPS.length > 1 ? 's' : '') + ' sin GPS instalado',
            accion: 'Ver flotilla',
            panel: '#panel-flotilla'
          });
        }
      }
    }
    
    // 2. Tickets urgentes (más de 3 días en Pendiente)
    var ticketsResp = obtenerTickets(token, {});
    if (ticketsResp.ok && ticketsResp.tickets) {
      var urgentes = ticketsResp.tickets.filter(function(t) {
        if (t.estado !== 'Pendiente') return false;
        try {
          var fecha = new Date(t.fecha);
          var diff = (new Date() - fecha) / (1000 * 60 * 60 * 24);
          return diff > 3;
        } catch(e) { return false; }
      });
      if (urgentes.length > 0) {
        alertas.push({
          tipo: 'danger',
          icono: '🚨',
          mensaje: urgentes.length + ' ticket' + (urgentes.length > 1 ? 's' : '') + ' sin resolver desde hace más de 3 días',
          accion: 'Ver tickets',
          panel: '#panel-tickets'
        });
      }
    }
    
    // 3. Borradores pendientes (solo técnico)
    if (!esAdminRevisor) {
      var borradoresResp = listarBorradores(token);
      if (borradoresResp.ok && borradoresResp.borradores) {
        var count = borradoresResp.borradores.length;
        if (count > 0) {
          alertas.push({
            tipo: 'warning',
            icono: '📝',
            mensaje: count + ' borradore' + (count > 1 ? 's' : '') + ' pendiente' + (count > 1 ? 's' : '') + ' de completar',
            accion: 'Ver borradores',
            panel: '#panel-borradores'
          });
        }
      }
    }
    
    // 4. Reportes pendientes de asignar tipo (solo Admin/Revisor)
    if (esAdminRevisor) {
      var registrosResp = obtenerRegistros(token, { estado: 'Listo para pago', limite: 100 });
      if (registrosResp.ok && registrosResp.registros) {
        var sinTipo = registrosResp.registros.filter(function(r) {
          return !r.TIPO_REVISION || r.TIPO_REVISION === 'PENDIENTE' || r.TIPO_REVISION === '';
        });
        if (sinTipo.length > 0) {
          alertas.push({
            tipo: 'info',
            icono: '📋',
            mensaje: sinTipo.length + ' reporte' + (sinTipo.length > 1 ? 's' : '') + ' pendiente' + (sinTipo.length > 1 ? 's' : '') + ' de asignar tipo',
            accion: 'Ver registros',
            panel: '#panel-registros'
          });
        }
      }
    }
    
    return { ok: true, alertas: alertas };
    
  } catch (err) {
    console.error('Error en obtenerAlertas:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Realiza copia de seguridad de todas las hojas
 * @param {string} token - Token de sesión (opcional, se usa desde trigger)
 * @returns {Object} { ok: true, mensaje: string, archivos: number }
 */
function realizarBackup(token) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var nombreArchivo = 'Backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    
    // Crear carpeta de backups (si no existe)
    var carpetaBackup = getCarpetaBackup();
    
    // Crear subcarpeta con fecha
    var fechaCarpeta = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var subCarpeta = carpetaBackup.createFolder(fechaCarpeta);
    
    // Copiar todas las hojas
    var hojas = ss.getSheets();
    var copiadas = 0;
    
    hojas.forEach(function(hoja) {
      var nombre = hoja.getName();
      // Saltar hojas de respaldo (no hacer backup de backups)
      if (nombre.startsWith('Backup_')) return;
      
      var datos = hoja.getDataRange().getValues();
      var blob = Utilities.newBlob(
        datos.map(function(row) { return row.join('\t'); }).join('\n'),
        'text/plain',
        nombre + '.txt'
      );
      subCarpeta.createFile(blob);
      copiadas++;
    });
    
    // Registrar en log
    var log = {
      fecha: new Date(),
      archivos: copiadas,
      carpeta: subCarpeta.getUrl()
    };
    
    // Guardar log en una hoja especial
    guardarLogBackup(log);
    
    // Limpiar backups antiguos (opcional)
    limpiarBackupsAntiguos(30); // Mantener solo últimos 30 días
    
    return {
      ok: true,
      mensaje: 'Backup completado',
      archivos: copiadas,
      carpeta: subCarpeta.getUrl()
    };
    
  } catch (err) {
    console.error('Error en backup:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Obtiene o crea la carpeta de backups
 */
function getCarpetaBackup() {
  var nombreCarpeta = 'Backups_FleetManager';
  var carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  
  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    return DriveApp.createFolder(nombreCarpeta);
  }
}

/**
 * Guarda el log del backup
 */
function guardarLogBackup(log) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('📈_Log_Auditoria');
  if (!sheet) return;
  
  sheet.appendRow([
    log.fecha,
    'BACKUP_AUTOMATICO',
    'SISTEMA',
    'Backup completado: ' + log.archivos + ' archivos',
    log.carpeta,
    '',
    ''
  ]);
}


/**
 * Configurar trigger diario (EJECUTAR UNA VEZ)
 * Ve a: Extensiones → Apps Script → Triggers (reloj) → Agregar trigger
 */
function configurarBackupAutomatico() {
  // Eliminar triggers existentes para evitar duplicados
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'realizarBackupProgramado') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Crear trigger diario a las 2:00 AM
  ScriptApp.newTrigger('realizarBackupProgramado')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
  
  console.log('✅ Backup automático configurado para las 2:00 AM diarias');
}

// ============================================================
// BACKUP AUTOMÁTICO - BACKEND
// ============================================================

/**
 * Ejecuta backup manual desde la interfaz
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, archivos: number, carpeta: string }
 */
function ejecutarBackupManual(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para ejecutar backup.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var carpetaBackup = obtenerCarpetaBackup();
    
    var fechaCarpeta = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    var subCarpeta = carpetaBackup.createFolder(fechaCarpeta);
    
    var hojas = ss.getSheets();
    var copiadas = 0;
    var errores = 0;

    hojas.forEach(function(hoja) {
      var nombre = hoja.getName();
      if (nombre.startsWith('Backup_')) return;
      
      try {
        var datos = hoja.getDataRange().getValues();
        var csv = datos.map(function(row) { 
          return row.map(function(cell) { 
            if (cell instanceof Date) {
              return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
            }
            if (typeof cell === 'string' && cell.includes(',')) {
              return '"' + cell + '"';
            }
            return cell; 
          }).join(','); 
        }).join('\n');
        
        var blob = Utilities.newBlob(csv, 'text/csv', nombre + '.csv');
        subCarpeta.createFile(blob);
        copiadas++;
      } catch(e) {
        errores++;
        console.error('Error copiando hoja:', nombre, e);
      }
    });

    // Guardar en auditoría
    _registrarAuditoria(
      sesionResp.sesion.usuarioId,
      sesionResp.sesion.nombre,
      'BACKUP',
      'SISTEMA',
      'Backup manual: ' + copiadas + ' archivos copiados' + (errores > 0 ? ' (' + errores + ' errores)' : ''),
      '',
      '',
      ''
    );

    // Guardar información del último backup en parámetros
    _escribirParam('BACKUP_ULTIMO', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
    _escribirParam('BACKUP_ARCHIVOS', String(copiadas));
    _escribirParam('BACKUP_CARPETA', subCarpeta.getUrl());

    // Limpiar backups antiguos
    var mantener = parseInt(_leerParams()['BACKUP_MANTENER'] || '30');
    limpiarBackupsAntiguos(mantener);

    return {
      ok: true,
      archivos: copiadas,
      carpeta: subCarpeta.getUrl(),
      errores: errores
    };

  } catch (err) {
    console.error('Error en backup manual:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// BACKUP AUTOMÁTICO
// ============================================================

/**
 * Obtiene la carpeta de backups en Drive
 * @returns {Folder} Carpeta de Drive
 */
function obtenerCarpetaBackup() {
  var nombreCarpeta = 'Backups_FleetManager';
  var carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  
  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    return DriveApp.createFolder(nombreCarpeta);
  }
}


/**
 * Obtiene la URL de la carpeta de backups
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, url: string }
 */
function obtenerUrlCarpetaBackup(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var carpeta = obtenerCarpetaBackup();
    return { ok: true, url: carpeta.getUrl() };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Limpia backups antiguos
 * @param {number} dias - Días a mantener
 */
function limpiarBackupsAntiguos(dias) {
  try {
    var carpeta = obtenerCarpetaBackup();
    var ahora = new Date();
    var fechaLimite = new Date(ahora);
    fechaLimite.setDate(fechaLimite.getDate() - dias);
    
    var subCarpetas = carpeta.getFolders();
    var eliminadas = 0;
    
    while (subCarpetas.hasNext()) {
      var sub = subCarpetas.next();
      var nombre = sub.getName();
      try {
        var fecha = new Date(nombre);
        if (fecha < fechaLimite) {
          sub.setTrashed(true);
          eliminadas++;
        }
      } catch(e) {
        sub.setTrashed(true);
        eliminadas++;
      }
    }
    
    console.log('🗑️ Eliminadas ' + eliminadas + ' carpetas de backup antiguas');
  } catch(e) {
    console.warn('Error limpiando backups antiguos:', e);
  }
}

/**
 * Carga información del último backup
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: true, ultimoBackup: string, archivos: number }
 */
function cargarInfoBackup(token) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  try {
    var params = _leerParams();
    var ultimo = params['BACKUP_ULTIMO'] || 'No hay backups registrados';
    var archivos = params['BACKUP_ARCHIVOS'] || '0';
    
    return {
      ok: true,
      ultimoBackup: ultimo,
      archivos: archivos
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Guarda la configuración de backup
 * @param {string} token - Token de sesión
 * @param {string} frecuencia - diario, semanal, mensual
 * @param {string} mantener - Días a mantener
 * @returns {Object} { ok: true }
 */
function guardarConfigBackup(token, frecuencia, mantener) {
  var sesionResp = validarSesion(token);
  if (!sesionResp.ok) return { ok: false, error: sesionResp.error };

  if (sesionResp.sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos.' };
  }

  try {
    _escribirParam('BACKUP_FRECUENCIA', frecuencia);
    _escribirParam('BACKUP_MANTENER', mantener);
    
    // Reconfigurar trigger automático
    configurarTriggerBackup(frecuencia);
    
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Configura el trigger automático para backup
 * @param {string} frecuencia - diario, semanal, mensual
 */
function configurarTriggerBackup(frecuencia) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'realizarBackupProgramado') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  var triggerBuilder = ScriptApp.newTrigger('realizarBackupProgramado').timeBased();
  
  switch(frecuencia) {
    case 'diario':
      triggerBuilder.atHour(2).everyDays(1);
      break;
    case 'semanal':
      triggerBuilder.onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3);
      break;
    case 'mensual':
      triggerBuilder.onMonthDay(1).atHour(4);
      break;
    default:
      triggerBuilder.atHour(2).everyDays(1);
  }
  
  triggerBuilder.create();
  console.log('✅ Trigger de backup configurado:', frecuencia);
}

/**
 * Función que ejecuta el backup desde el trigger (NO requiere token)
 */
function realizarBackupProgramado() {
  var resultado = ejecutarBackupManual(null);
  
  if (!resultado.ok) {
    try {
      var email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: '⚠️ ALERTA: Falló el backup automático',
          body: 'El backup automático falló:\n\n' + resultado.error
        });
      }
    } catch(e) {}
  }
  
  return resultado;
}


/**
 * Configurar backup inicial (ejecutar UNA VEZ en el editor)
 */
function configurarBackupInicial() {
  var params = _leerParams();
  var frecuencia = params['BACKUP_FRECUENCIA'] || 'diario';
  configurarTriggerBackup(frecuencia);
  console.log('✅ Backup automático configurado en:', frecuencia);
}
function verificarColumnas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var hojas = {
    '📦_Inventario_GPS': 'ECONOMICO_ASIGNADO',
    '🔧_Accesorios_Stock': 'ECONOMICO_ASIGNADO',
    '📋_Catalogo_Vehiculos': 'ECONOMICO',
    '🚚_Flotilla_Fallas': 'ECONOMICO',
    '📝_Bitacora_Revisiones': 'ECONOMICO'
  };
  
  for (var nombreHoja in hojas) {
    var sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) {
      console.log('❌ No existe:', nombreHoja);
      continue;
    }
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnaBuscada = hojas[nombreHoja];
    var existe = headers.indexOf(columnaBuscada) !== -1;
    console.log(nombreHoja + ' → ' + columnaBuscada + ': ' + (existe ? '✅' : '❌'));
  }
}
/**
 * Guarda los cambios de un ticket editado
 * @param {string} token - Token de sesión
 * @param {string} ticketId - ID del ticket
 * @param {Object} datos - { unidad, descripcion, estado, prioridad, comentarios, notasInternas }
 * @returns {Object} { ok, mensaje, error }
 */
function actualizarTicket(token, ticketId, datos) {
  try {
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      return { ok: false, error: sesionResp.error };
    }

    const sesion = sesionResp.sesion;
    const sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Tickets' };
    }

    const datosSheet = sheet.getDataRange().getValues();
    const headers = datosSheet[0];

    // Mapear índices
    const idxId = headers.indexOf('ID');
    const idxUnidad = headers.indexOf('UNIDAD');
    const idxDescripcion = headers.indexOf('DESCRIPCION');
    const idxEstado = headers.indexOf('ESTADO');
    const idxPrioridad = headers.indexOf('PRIORIDAD');
    const idxComentarios = headers.indexOf('COMENTARIOS');
    const idxNotasInternas = headers.indexOf('NOTAS_INTERNAS');
    const idxUltimaActualizacion = headers.indexOf('ULTIMA_ACTUALIZACION');
    const idxCreadoPor = headers.indexOf('CREADO_POR');

    if (idxId === -1) {
      return { ok: false, error: 'Estructura de tickets incompleta' };
    }

    // Buscar el ticket
    let filaEncontrada = -1;
    for (var i = 1; i < datosSheet.length; i++) {
      if (datosSheet[i][idxId] === ticketId) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { ok: false, error: 'Ticket no encontrado: ' + ticketId };
    }

    // ✅ Verificar permisos
    const esAdmin = sesion.rol === 1 || sesion.rol === 2;
    if (!esAdmin && datosSheet[filaEncontrada - 1][idxCreadoPor] !== sesion.usuarioId) {
      return { ok: false, error: 'No tienes permisos para editar este ticket.' };
    }

    // Actualizar campos
    if (idxUnidad !== -1 && datos.unidad) {
      sheet.getRange(filaEncontrada, idxUnidad + 1).setValue(datos.unidad);
    }
    if (idxDescripcion !== -1 && datos.descripcion) {
      sheet.getRange(filaEncontrada, idxDescripcion + 1).setValue(datos.descripcion);
    }
    if (idxEstado !== -1 && datos.estado) {
      sheet.getRange(filaEncontrada, idxEstado + 1).setValue(datos.estado);
    }
    if (idxPrioridad !== -1 && datos.prioridad) {
      sheet.getRange(filaEncontrada, idxPrioridad + 1).setValue(datos.prioridad);
    }
    if (idxComentarios !== -1) {
      const comentariosActuales = datosSheet[filaEncontrada - 1][idxComentarios] || '';
      const nuevoComentario = comentariosActuales 
        ? comentariosActuales + '\n--- EDICIÓN ---\n' + datos.comentarios + '\n' + sesion.nombre + ' - ' + new Date().toLocaleString()
        : 'EDICIÓN:\n' + datos.comentarios + '\n' + sesion.nombre + ' - ' + new Date().toLocaleString();
      sheet.getRange(filaEncontrada, idxComentarios + 1).setValue(nuevoComentario);
    }
    if (idxNotasInternas !== -1) {
      sheet.getRange(filaEncontrada, idxNotasInternas + 1).setValue(datos.notasInternas || '');
    }
    if (idxUltimaActualizacion !== -1) {
      sheet.getRange(filaEncontrada, idxUltimaActualizacion + 1).setValue(new Date());
    }

    // Registrar en auditoría
    _registrarAuditoria(
      sesion.usuarioId,
      sesion.nombre,
      'EDITAR_TICKET',
      'TICKETS',
      'Ticket ' + ticketId + ' editado por ' + sesion.nombre,
      ticketId,
      '',
      ''
    );

    return { ok: true, mensaje: 'Ticket actualizado correctamente' };

  } catch (err) {
    console.error('❌ Error en actualizarTicket:', err);
    return { ok: false, error: err.message };
  }
}
/**
 * Obtiene la lista de vehículos de la flotilla para el selector de tickets
 * @param {string} token - Token de sesión
 * @returns {Object} { ok: boolean, vehiculos: [{economico, placas}] }
 */
function obtenerVehiculosFlotilla(token) {
  try {
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      return { ok: false, error: sesionResp.error };
    }

    const sheet = SHEETS.FLOTILLA();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Flotilla' };
    }

    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];

    const idxEconomico = headers.indexOf('ECONOMICO');
    const idxPlacas = headers.indexOf('PLACAS');

    if (idxEconomico === -1) {
      return { ok: false, error: 'Estructura de flotilla incompleta: falta columna ECONOMICO' };
    }

    const vehiculos = [];
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxEconomico]) {
        vehiculos.push({
          economico: datos[i][idxEconomico].toString(),
          placas: datos[i][idxPlacas] || ''
        });
      }
    }

    return { ok: true, vehiculos: vehiculos };

  } catch (err) {
    console.error('❌ Error en obtenerVehiculosFlotilla:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// PRUEBA 1: Subir fotos a carpeta existente
// ============================================================
function testAnexarFotos() {
  // Crear carpeta de prueba
  const carpetaPrueba = DriveApp.createFolder('Test_Fotos_' + Date.now());
  const folderUrl = carpetaPrueba.getUrl();
  
  // Crear archivos de prueba (imágenes pequeñas en Base64)
  const archivosPrueba = [
    {
      nombre: 'test_1.jpg',
      base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      mimeType: 'image/jpeg'
    },
    {
      nombre: 'test_2.jpg',
      base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      mimeType: 'image/jpeg'
    }
  ];
  
  // Ejecutar función
  const resultado = _anexarFotosACarpetaExistente(folderUrl, archivosPrueba);
  
  // Mostrar resultado
  console.log('Resultado:', resultado);
  
  // Verificar
  if (resultado.success) {
    console.log('✅ Prueba exitosa: ' + resultado.archivosSubidos + ' archivos subidos');
  } else {
    console.error('❌ Prueba fallida:', resultado.errores);
  }
  
  // Limpiar carpeta de prueba
  // carpetaPrueba.setTrashed(true);
}

// ============================================================
// PRUEBA 2: Manejo de errores
// ============================================================
function testAnexarFotosErrores() {
  // Caso 1: URL inválida
  const resultado1 = _anexarFotosACarpetaExistente('URL_INVALIDA', [{ nombre: 'test.jpg', base64: 'data' }]);
  console.log('Caso 1 (URL inválida):', resultado1);
  
  // Caso 2: Sin archivos
  const carpeta = DriveApp.createFolder('Test_Error_' + Date.now());
  const resultado2 = _anexarFotosACarpetaExistente(carpeta.getUrl(), []);
  console.log('Caso 2 (Sin archivos):', resultado2);
  carpeta.setTrashed(true);
  
  // Caso 3: Archivo corrupto
  const resultado3 = _anexarFotosACarpetaExistente(carpeta.getUrl(), [
    { nombre: 'corrupto.jpg', base64: 'DATOS_INVALIDOS' }
  ]);
  console.log('Caso 3 (Archivo corrupto):', resultado3);
}
// ============================================================
// 12. GESTIÓN DE TICKETS - ACTUALIZAR ESTADO
// ============================================================

/**
 * Actualiza el estado de un ticket en la hoja 🎫_Tickets
 * @param {string} token - Token de sesión
 * @param {string} ticketId - ID del ticket (ej. TKT-002)
 * @param {string} nuevoEstado - Nuevo estado (Pendiente, En proceso, Resuelto)
 * @returns {Object} { ok, mensaje, error }
 */
function actualizarEstadoTicket(token, ticketId, nuevoEstado) {
  // 1. Validar sesión
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }

  const sesion = sesionResp.sesion;

  // 2. Validar que el estado sea válido
  const estadosValidos = ['Pendiente', 'En proceso', 'Resuelto'];
  if (!estadosValidos.includes(nuevoEstado)) {
    return { ok: false, error: 'Estado inválido. Usa: Pendiente, En proceso, Resuelto' };
  }

  try {
    const sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Tickets' };
    }

    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];

    // Buscar índices
    const idxId = headers.indexOf('ID');
    const idxEstado = headers.indexOf('ESTADO');
    const idxTecnico = headers.indexOf('TECNICO_ASIGNADO');
    const idxFechaCierre = headers.indexOf('FECHA_CIERRE');

    if (idxId === -1 || idxEstado === -1) {
      return { ok: false, error: 'Columnas ID o ESTADO no encontradas' };
    }

    // Buscar el ticket
    let filaEncontrada = -1;

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxId] === ticketId) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { ok: false, error: 'Ticket no encontrado: ' + ticketId };
    }

    // 3. Actualizar estado
    sheet.getRange(filaEncontrada, idxEstado + 1).setValue(nuevoEstado);

    // 4. Si se resuelve, registrar fecha de cierre
    if (nuevoEstado === 'Resuelto' && idxFechaCierre !== -1) {
      sheet.getRange(filaEncontrada, idxFechaCierre + 1).setValue(new Date());
    }

    // 5. Registrar en auditoría
    _registrarAuditoria(
      sesion.usuarioId,
      sesion.nombre,
      'ACTUALIZAR_TICKET',
      'TICKETS',
      'Ticket ' + ticketId + ' cambiado a "' + nuevoEstado + '" por ' + sesion.nombre,
      ticketId,
      '',
      ''
    );

    // 6. Crear notificación para el técnico asignado (si existe)
    if (idxTecnico !== -1) {
      const tecnicoId = datos[filaEncontrada - 1][idxTecnico];
      if (tecnicoId) {
        _crearNotificacion(
          tecnicoId,
          'TICKET_ACTUALIZADO',
          '📌 Ticket ' + ticketId + ' ha cambiado a "' + nuevoEstado + '"',
          ticketId,
          '#panel-tickets'
        );
      }
    }

    console.log('✅ Ticket ' + ticketId + ' actualizado a: ' + nuevoEstado);
    return {
      ok: true,
      mensaje: 'Ticket ' + ticketId + ' actualizado a "' + nuevoEstado + '"'
    };

  } catch (err) {
    console.error('❌ Error al actualizar ticket:', err);
    return { ok: false, error: 'Error al actualizar ticket: ' + err.message };
  }
}
// ============================================================
// TICKETS - GESTIÓN DE TÉCNICOS AUTORIZADOS
// ============================================================

/**
 * Obtiene la lista de técnicos y los autorizados actuales de un ticket
 * @param {string} token - Token de sesión
 * @param {string} ticketId - ID del ticket
 * @returns {Object} { ok, tecnicos: [], autorizados: [] }
 */
function obtenerTecnicosYAutorizados(token, ticketId) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }

  const sesion = sesionResp.sesion;

  // Solo Admin/Revisor pueden gestionar
  if (sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para gestionar técnicos.' };
  }

  try {
    // 1. Obtener todos los técnicos activos
    const usuariosSheet = SHEETS.USUARIOS();
    const usuariosData = usuariosSheet.getDataRange().getValues();
    const tecnicos = [];

    for (var i = 1; i < usuariosData.length; i++) {
      const rol = usuariosData[i][4];
      const activo = usuariosData[i][5];
      if (rol === 3 && activo === true) {
        tecnicos.push({
          id: usuariosData[i][0].toString(),
          nombre: usuariosData[i][1].toString()
        });
      }
    }

    // 2. Obtener autorizados actuales del ticket
    const ticketsSheet = SHEETS.TICKETS();
    const ticketsData = ticketsSheet.getDataRange().getValues();
    const headers = ticketsData[0];
    const idxId = headers.indexOf('ID');
    const idxTecnicosAutorizados = headers.indexOf('TECNICOS_AUTORIZADOS');

    let autorizados = [];
    for (var j = 1; j < ticketsData.length; j++) {
      if (ticketsData[j][idxId] === ticketId) {
        const autorizadosStr = ticketsData[j][idxTecnicosAutorizados] || '';
        autorizados = autorizadosStr ? autorizadosStr.split(',') : [];
        break;
      }
    }

    return {
      ok: true,
      tecnicos: tecnicos,
      autorizados: autorizados
    };

  } catch (err) {
    console.error('❌ Error en obtenerTecnicosYAutorizados:', err);
    return { ok: false, error: err.message };
  }
}

/**
 * Guarda la lista de técnicos autorizados para un ticket
 * @param {string} token - Token de sesión
 * @param {string} ticketId - ID del ticket
 * @param {Array} autorizados - Lista de IDs de técnicos autorizados
 * @returns {Object} { ok, mensaje }
 */
function guardarTecnicosAutorizados(token, ticketId, autorizados) {
  const sesionResp = validarSesion(token);
  if (!sesionResp.ok) {
    return { ok: false, error: sesionResp.error };
  }

  const sesion = sesionResp.sesion;

  if (sesion.rol > 2) {
    return { ok: false, error: 'Sin permisos para gestionar técnicos.' };
  }

  try {
    const sheet = SHEETS.TICKETS();
    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];
    const idxId = headers.indexOf('ID');
    const idxTecnicosAutorizados = headers.indexOf('TECNICOS_AUTORIZADOS');

    if (idxId === -1 || idxTecnicosAutorizados === -1) {
      return { ok: false, error: 'Estructura de tickets incompleta.' };
    }

    let filaEncontrada = -1;
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxId] === ticketId) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) {
      return { ok: false, error: 'Ticket no encontrado: ' + ticketId };
    }

    const autorizadosStr = autorizados.join(',');
    sheet.getRange(filaEncontrada, idxTecnicosAutorizados + 1).setValue(autorizadosStr);

    _registrarAuditoria(
      sesion.usuarioId,
      sesion.nombre,
      'EDITAR_TICKET',
      'TICKETS',
      'Técnicos autorizados actualizados para ticket ' + ticketId + ': ' + autorizadosStr,
      ticketId,
      '',
      ''
    );

    return {
      ok: true,
      mensaje: 'Técnicos autorizados actualizados correctamente',
      autorizados: autorizados
    };

  } catch (err) {
    console.error('❌ Error en guardarTecnicosAutorizados:', err);
    return { ok: false, error: err.message };
  }
}
// ============================================================
// TICKETS - OBTENER TICKET POR ID
// ============================================================

/**
 * Obtiene un ticket específico por su ID (con control de permisos)
 * @param {string} token - Token de sesión
 * @param {string} ticketId - ID del ticket
 * @returns {Object} { ok, ticket, error }
 */
function obtenerTicketPorId(token, ticketId) {
  try {
    const sesionResp = validarSesion(token);
    if (!sesionResp.ok) {
      return { ok: false, error: sesionResp.error };
    }

    const sesion = sesionResp.sesion;
    const sheet = SHEETS.TICKETS();
    if (!sheet) {
      return { ok: false, error: 'No se encontró la hoja de Tickets' };
    }

    const datos = sheet.getDataRange().getValues();
    const headers = datos[0];

    // Mapear índices
    const idxId = headers.indexOf('ID');
    const idxFecha = headers.indexOf('FECHA');
    const idxUnidad = headers.indexOf('UNIDAD');
    const idxDescripcion = headers.indexOf('DESCRIPCION');
    const idxCreadoPor = headers.indexOf('CREADO_POR');
    const idxCreadoPorNombre = headers.indexOf('CREADO_POR_NOMBRE');
    const idxEstado = headers.indexOf('ESTADO');
    const idxTecnicoAsignado = headers.indexOf('TECNICO_ASIGNADO');
    const idxTecnicoNombre = headers.indexOf('TECNICO_NOMBRE');
    const idxFechaCierre = headers.indexOf('FECHA_CIERRE');
    const idxComentarios = headers.indexOf('COMENTARIOS');
    const idxUltimaActualizacion = headers.indexOf('ULTIMA_ACTUALIZACION');
    const idxTecnicosAutorizados = headers.indexOf('TECNICOS_AUTORIZADOS');

    if (idxId === -1) {
      return { ok: false, error: 'Estructura de tickets incompleta' };
    }

    const esAdmin = sesion.rol === 1 || sesion.rol === 2;
    const usuarioId = sesion.usuarioId;

    // Buscar el ticket
    let fila = null;
    let filaIndex = -1;

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][idxId] === ticketId) {
        fila = datos[i];
        filaIndex = i;
        break;
      }
    }

    if (!fila) {
      return { ok: false, error: 'Ticket no encontrado: ' + ticketId };
    }

    // ✅ CONTROL DE VISIBILIDAD
    let puedeVer = false;

    if (esAdmin) {
      puedeVer = true;
    } else if (fila[idxCreadoPor] === usuarioId) {
      puedeVer = true;
    } else if (idxTecnicosAutorizados !== -1) {
      const autorizados = fila[idxTecnicosAutorizados]
        ? fila[idxTecnicosAutorizados].toString().split(',')
        : [];
      puedeVer = autorizados.includes(usuarioId);
    } else if (fila[idxTecnicoAsignado] === usuarioId) {
      puedeVer = true;
    }

    if (!puedeVer) {
      return { ok: false, error: 'No tienes permisos para ver este ticket.' };
    }

    // Construir objeto ticket
    const ticket = {
      id: fila[idxId] || '',
      fecha: fila[idxFecha] ? _formatearFecha(fila[idxFecha]) : '—',
      unidad: fila[idxUnidad] || '—',
      descripcion: fila[idxDescripcion] || '—',
      creadoPor: fila[idxCreadoPor] || '',
      creadoPorNombre: fila[idxCreadoPorNombre] || fila[idxCreadoPor] || '—',
      estado: fila[idxEstado] || 'Pendiente',
      tecnico: fila[idxTecnicoAsignado] || '',
      tecnicoNombre: fila[idxTecnicoNombre] || '—',
      fechaCierre: fila[idxFechaCierre] ? _formatearFecha(fila[idxFechaCierre]) : null,
      comentarios: fila[idxComentarios] || '',
      ultimaActualizacion: fila[idxUltimaActualizacion] ? _formatearFecha(fila[idxUltimaActualizacion]) : '—',
      tecnicosAutorizados: fila[idxTecnicosAutorizados] || ''
    };

    return { ok: true, ticket: ticket };

  } catch (err) {
    console.error('❌ Error en obtenerTicketPorId:', err);
    return { ok: false, error: 'Error al obtener ticket: ' + err.message };
  }
}