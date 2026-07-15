// ============================================================
// SISTEMA DE GESTIÓN DE FLOTAS — VerificacionBaseDatos.gs
// Verifica la integridad de la base de datos y permite instalarla
// desde la interfaz web.
// ============================================================

/**
 * Verifica si la base de datos está completamente instalada.
 * @returns {Object} { ok: boolean, faltantes: string[], instalado: boolean }
 */
function verificarBaseDatos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojasRequeridas = [
      '💰_Tarifas',
      '👤_Usuarios',
      '⚙️_Parametros',
      '📦_Inventario_GPS',
      '🚚_Flotilla_Fallas',
      '📝_Bitacora_Revisiones',
      '📊_Consulta_Tecnicos',
      '📋_Tipos_Equipo',
      '🔧_Accesorios_Stock',
      '📑_Facturas',
      '📈_Log_Auditoria',
      '📩_Notificaciones'
    ];
    
    var hojasExistentes = ss.getSheets().map(function(s) { return s.getName(); });
    var faltantes = [];
    
    for (var i = 0; i < hojasRequeridas.length; i++) {
      if (hojasExistentes.indexOf(hojasRequeridas[i]) === -1) {
        faltantes.push(hojasRequeridas[i]);
      }
    }
    
    // Verificar que las hojas tengan encabezados
    var estructuraValida = true;
    if (faltantes.length === 0) {
      estructuraValida = verificarEstructuraHojas(ss);
    }
    
    return {
      ok: faltantes.length === 0 && estructuraValida,
      instalado: faltantes.length === 0 && estructuraValida,
      faltantes: faltantes,
      mensaje: faltantes.length > 0 
        ? 'Faltan ' + faltantes.length + ' hoja(s) por instalar.'
        : (estructuraValida ? 'Base de datos completa.' : 'Estructura de hojas incompleta.')
    };
    
  } catch (err) {
    console.error('Error en verificarBaseDatos:', err);
    return {
      ok: false,
      instalado: false,
      faltantes: [],
      mensaje: 'Error al verificar: ' + err.message,
      error: err.message
    };
  }
}

/**
 * Verifica que las hojas tengan los encabezados correctos.
 */
function verificarEstructuraHojas(ss) {
  try {
    // Verificar tarifas
    var tarifas = ss.getSheetByName('💰_Tarifas');
    if (tarifas) {
      var headers = tarifas.getRange(1, 1, 1, 5).getValues()[0];
      var esperados = ['TIPO', 'DESCRIPCION', 'PRECIO', 'CLAVE_INTERNA', 'ACTIVO'];
      for (var i = 0; i < esperados.length; i++) {
        if (headers[i] !== esperados[i]) return false;
      }
    } else {
      return false;
    }
    
    // Verificar usuarios
    var usuarios = ss.getSheetByName('👤_Usuarios');
    if (usuarios) {
      var headers = usuarios.getRange(1, 1, 1, 7).getValues()[0];
      var esperados = ['ID_USUARIO', 'NOMBRE', 'USUARIO', 'PASS_HASH', 'ROL', 'ACTIVO', 'ULTIMO_ACCESO'];
      for (var i = 0; i < esperados.length; i++) {
        if (headers[i] !== esperados[i]) return false;
      }
    } else {
      return false;
    }
    
    return true;
    
  } catch (err) {
    console.error('Error en verificarEstructuraHojas:', err);
    return false;
  }
}

/**
 * Instala la base de datos desde la web app.
 * Esta función llama al script de instalación.
 */
function instalarBaseDatosDesdeWeb() {
  try {
    console.log('=== INICIO instalarBaseDatosDesdeWeb ===');
    
    // Verificar si ya está instalada
    var verificacion = verificarBaseDatos();
    console.log('Verificación pre-instalación:', verificacion);
    
    if (verificacion.instalado) {
      console.log('✅ Ya está instalada');
      return {
        ok: false,
        mensaje: 'La base de datos ya está instalada correctamente.'
      };
    }
    
    console.log('🔄 Ejecutando instalación...');
    // Ejecutar instalación
    instalarBaseDatos();
    
    // Verificar que la instalación fue exitosa
    var verificacionPost = verificarBaseDatos();
    console.log('Verificación post-instalación:', verificacionPost);
    
    if (verificacionPost.instalado) {
      console.log('✅ Instalación exitosa');
      return {
        ok: true,
        mensaje: 'Base de datos instalada exitosamente.'
      };
    } else {
      console.warn('⚠️ Instalación incompleta');
      return {
        ok: false,
        mensaje: 'La instalación no se completó correctamente. Faltan: ' + 
                 verificacionPost.faltantes.join(', ')
      };
    }
    
  } catch (err) {
    console.error('❌ Error en instalarBaseDatosDesdeWeb:', err);
    console.error('Stack:', err.stack);
    return {
      ok: false,
      mensaje: 'Error durante la instalación: ' + err.message
    };
  }
}

/**
 * Obtiene el estado de la base de datos para mostrar en la interfaz.
 * Incluye información detallada para el administrador.
 */
function obtenerEstadoBaseDatos() {
  try {
    console.log('=== INICIO obtenerEstadoBaseDatos ===');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var verificacion = verificarBaseDatos();
    
    var hojasExistentes = ss.getSheets().map(function(s) { return s.getName(); });
    console.log('Hojas existentes:', hojasExistentes.join(', '));
    
    var info = {
      ok: verificacion.ok,
      instalado: verificacion.instalado,
      faltantes: verificacion.faltantes,
      hojasExistentes: hojasExistentes,
      totalHojas: hojasExistentes.length,
      mensaje: verificacion.mensaje,
      detalles: {
        tarifas: hojasExistentes.indexOf('💰_Tarifas') !== -1,
        usuarios: hojasExistentes.indexOf('👤_Usuarios') !== -1,
        parametros: hojasExistentes.indexOf('⚙️_Parametros') !== -1,
        inventario: hojasExistentes.indexOf('📦_Inventario_GPS') !== -1,
        flotilla: hojasExistentes.indexOf('🚚_Flotilla_Fallas') !== -1,
        bitacora: hojasExistentes.indexOf('📝_Bitacora_Revisiones') !== -1,
        consulta: hojasExistentes.indexOf('📊_Consulta_Tecnicos') !== -1,
        tiposEquipo: hojasExistentes.indexOf('📋_Tipos_Equipo') !== -1,
        accesorios: hojasExistentes.indexOf('🔧_Accesorios_Stock') !== -1,
        facturas: hojasExistentes.indexOf('📑_Facturas') !== -1,
        logAuditoria: hojasExistentes.indexOf('📈_Log_Auditoria') !== -1,
        notificaciones: hojasExistentes.indexOf('📩_Notificaciones') !== -1
      }
    };
    
    console.log('Estado DB - instalado:', info.instalado, 'faltantes:', info.faltantes);
    return info;
    
  } catch (err) {
    console.error('❌ Error en obtenerEstadoBaseDatos:', err);
    console.error('Stack:', err.stack);
    return {
      ok: false,
      instalado: false,
      mensaje: 'Error al obtener estado: ' + err.message
    };
  }
}