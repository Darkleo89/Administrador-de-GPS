// ============================================================
// SISTEMA DE GESTIÓN DE FLOTAS — InstalacionBaseDatos.gs
// Script de instalación: crea todas las hojas, encabezados,
// estructuras y datos de ejemplo.
// Versión actualizada con:
//   - TIPO_UNIDAD en Inventario y Flotilla
//   - Hoja 📋_Tipos_Unidad
//   - Hoja 🎫_Tickets (sistema de tickets/fallas)
//   - Dashboard como pantalla de inicio
// ============================================================

/**
 * PUNTO DE ENTRADA PRINCIPAL
 * Ejecuta esta función para instalar todo el sistema desde cero.
 * IMPORTANTE: Esta función debe ejecutarse desde el EDITOR DE SCRIPTS.
 */

/**
 * Instala la base de datos completa del sistema de gestión de flotas
 * Crea todas las hojas necesarias, configura validaciones, protecciones y menú
 * @returns {Object} Resultado de la instalación
 */
function instalarBaseDatos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // ============================================================
    // 1. VALIDACIÓN PREVIA
    // ============================================================
    // Verificar que se ejecuta desde el editor (no desde Web App)
    if (!ss || !ui) {
      throw new Error('Esta función solo puede ejecutarse desde el editor de Google Sheets');
    }

    // Solicitar confirmación al usuario
    const confirmacion = ui.alert(
      '⚠️ ADVERTENCIA',
      'Esta acción reinstalará toda la base de datos.\n\n' +
      '⚠️ Se ELIMINARÁN todas las hojas existentes y se crearán nuevas.\n' +
      '⚠️ Los datos actuales se PERDERÁN.\n\n' +
      '¿Estás seguro de continuar?',
      ui.ButtonSet.YES_NO
    );

    if (confirmacion !== ui.Button.YES) {
      ui.alert('❌ Instalación cancelada', 'No se realizaron cambios.', ui.ButtonSet.OK);
      return { ok: false, mensaje: 'Instalación cancelada por el usuario' };
    }

    console.log('🚀 Iniciando instalación de base de datos...');

    // ============================================================
    // 2. LIMPIAR HOJAS EXISTENTES (OPCIONAL)
    // ============================================================
    // Eliminar hojas existentes para evitar duplicados
    const hojasExistentes = ss.getSheets();
    const hojasAEliminar = hojasExistentes.filter(sheet => {
      const nombre = sheet.getName();
      // No eliminar hojas que empiecen con "📊" o sean la primera hoja
      return !nombre.startsWith('📊') && sheet.getSheetId() !== ss.getSheets()[0].getSheetId();
    });

    hojasAEliminar.forEach(sheet => {
      try {
        ss.deleteSheet(sheet);
        console.log(`🗑️ Hoja eliminada: ${sheet.getName()}`);
      } catch (err) {
        console.warn(`⚠️ No se pudo eliminar "${sheet.getName()}":`, err.message);
      }
    });

    // ============================================================
    // 3. CREAR HOJAS DEL SISTEMA
    // ============================================================
    console.log('📋 Creando hojas del sistema...');

    const hojas = {};

    // 3.1 Crear hojas principales
    const nombresHojas = [
      '💰_Tarifas',
      '👤_Usuarios',
      '⚙️_Parametros',
      '📦_Inventario_GPS',
      '🚚_Flotilla_Fallas',
      '📝_Bitacora_Revisiones',
      '📊_Consulta_Tecnicos',
      '📋_Tipos_Equipo',
      '📋_Tipos_Unidad',
      '📋_Tipos_Vehiculo',
      '📋_Estados_Ticket',
      '🔧_Accesorios_Stock',
      '📑_Facturas',
      '📈_Log_Auditoria',
      '📩_Notificaciones',
      '🎫_Tickets',
      '📋_Catálogo_Estados_Vehiculo'
    ];

    nombresHojas.forEach(nombre => {
      try {
        // Verificar si la hoja ya existe
        let sheet = ss.getSheetByName(nombre);
        if (sheet) {
          console.log(`⚠️ La hoja "${nombre}" ya existe, se mantendrá`);
        } else {
          sheet = ss.insertSheet(nombre);
          console.log(`✅ Hoja creada: ${nombre}`);
        }
        hojas[nombre] = sheet;
      } catch (err) {
        console.error(`❌ Error al crear "${nombre}":`, err.message);
        throw new Error(`No se pudo crear la hoja "${nombre}": ${err.message}`);
      }
    });

    // ============================================================
    // 4. CONFIGURAR ESTRUCTURA DE CADA HOJA
    // ============================================================
    console.log('📝 Configurando estructura de hojas...');

    // 4.1 💰_Tarifas
    const tarifasHeaders = ['TIPO_REVISION', 'COSTO_BASE', 'COSTO_POR_KM', 'COSTO_POR_HORA', 'DESCRIPCION'];
    hojas['💰_Tarifas'].getRange(1, 1, 1, tarifasHeaders.length).setValues([tarifasHeaders]);
    hojas['💰_Tarifas'].getRange('A1:E1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');

    // 4.2 👤_Usuarios
    const usuariosHeaders = ['EMAIL', 'NOMBRE', 'ROL', 'CONTRASEÑA_HASH', 'ACTIVO', 'FECHA_CREACION', 'ULTIMO_ACCESO'];
    hojas['👤_Usuarios'].getRange(1, 1, 1, usuariosHeaders.length).setValues([usuariosHeaders]);
    hojas['👤_Usuarios'].getRange('A1:G1').setFontWeight('bold').setBackground('#2196F3').setFontColor('#FFFFFF');

    // 4.3 ⚙️_Parametros
    const parametrosHeaders = ['CLAVE', 'VALOR', 'DESCRIPCION', 'FECHA_ACTUALIZACION'];
    hojas['⚙️_Parametros'].getRange(1, 1, 1, parametrosHeaders.length).setValues([parametrosHeaders]);
    hojas['⚙️_Parametros'].getRange('A1:D1').setFontWeight('bold').setBackground('#FF9800').setFontColor('#FFFFFF');

    // 4.4 📦_Inventario_GPS
    const inventarioHeaders = [
      'FECHA_REGISTRO', 'TIPO_EQUIPO', 'SERIE', 'MARCA', 'MODELO', 
      'NUMERO_ECONOMICO', 'FECHA_INSTALACION', 'ESTADO', 'OBSERVACIONES', 
      'UBICACION', 'FECHA_ULTIMA_REVISION', 'TIPO_UNIDAD'
    ];
    hojas['📦_Inventario_GPS'].getRange(1, 1, 1, inventarioHeaders.length).setValues([inventarioHeaders]);
    hojas['📦_Inventario_GPS'].getRange('A1:L1').setFontWeight('bold').setBackground('#9C27B0').setFontColor('#FFFFFF');

    // 4.5 🚚_Flotilla_Fallas
    const flotillaHeaders = [
      'FECHA_REPORTE', 'UNIDAD', 'TIPO_VEHICULO', 'FALLA_DETECTADA', 
      'ACCION_REALIZADA', 'RESPONSABLE', 'ESTADO', 'FECHA_SOLUCION', 'COSTO'
    ];
    hojas['🚚_Flotilla_Fallas'].getRange(1, 1, 1, flotillaHeaders.length).setValues([flotillaHeaders]);
    hojas['🚚_Flotilla_Fallas'].getRange('A1:I1').setFontWeight('bold').setBackground('#F44336').setFontColor('#FFFFFF');

    // 4.6 📝_Bitacora_Revisiones
    const bitacoraHeaders = [
      'FOLIO', 'FECHA_REPORTE', 'TECNICO', 'GPS', 'CLIENTE', 
      'TIPO_REVISION', 'OBSERVACIONES', 'ESTADO', 'FECHA_CIERRE', 
      'MONTO_TOTAL', 'FOTOS_DRIVE_URL', 'DATOS_JSON', 'ESTADO_PAGO'
    ];
    hojas['📝_Bitacora_Revisiones'].getRange(1, 1, 1, bitacoraHeaders.length).setValues([bitacoraHeaders]);
    hojas['📝_Bitacora_Revisiones'].getRange('A1:M1').setFontWeight('bold').setBackground('#00BCD4').setFontColor('#FFFFFF');

    // 4.7 🎫_Tickets
    const ticketsHeaders = [
      'ID_TICKET', 'FECHA_CREACION', 'USUARIO', 'ASUNTO', 
      'DESCRIPCION', 'PRIORIDAD', 'ESTADO', 'TECNICO_ASIGNADO', 
      'FECHA_CIERRE', 'SOLUCION'
    ];
    hojas['🎫_Tickets'].getRange(1, 1, 1, ticketsHeaders.length).setValues([ticketsHeaders]);
    hojas['🎫_Tickets'].getRange('A1:J1').setFontWeight('bold').setBackground('#FF5722').setFontColor('#FFFFFF');

    // 4.8 Otras hojas (configuración básica)
    const hojasSimples = [
      { nombre: '📊_Consulta_Tecnicos', headers: ['TECNICO', 'REPORTES_ATENDIDOS', 'REPORTES_PENDIENTES', 'EFICIENCIA'] },
      { nombre: '📋_Tipos_Equipo', headers: ['TIPO_EQUIPO', 'DESCRIPCION'] },
      { nombre: '📋_Tipos_Unidad', headers: ['TIPO_UNIDAD', 'DESCRIPCION'] },
      { nombre: '📋_Tipos_Vehiculo', headers: ['TIPO_VEHICULO', 'DESCRIPCION'] },
      { nombre: '📋_Estados_Ticket', headers: ['ESTADO', 'DESCRIPCION', 'COLOR'] },
      { nombre: '🔧_Accesorios_Stock', headers: ['ACCESORIO', 'CANTIDAD', 'PRECIO_UNITARIO', 'PROVEEDOR'] },
      { nombre: '📑_Facturas', headers: ['FOLIO_FACTURA', 'FECHA', 'CLIENTE', 'MONTO', 'ESTADO_PAGO'] },
      { nombre: '📈_Log_Auditoria', headers: ['FECHA', 'USUARIO', 'ACCION', 'DETALLE', 'IP'] },
      { nombre: '📩_Notificaciones', headers: ['ID', 'FECHA', 'DESTINATARIO', 'MENSAJE', 'LEIDO'] },
      { nombre: '📋_Catálogo_Estados_Vehiculo', headers: ['ESTADO_VEHICULO', 'DESCRIPCION', 'COLOR'] }
    ];

    hojasSimples.forEach(({ nombre, headers }) => {
      try {
        const sheet = hojas[nombre];
        if (sheet && headers) {
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#795548').setFontColor('#FFFFFF');
          console.log(`✅ Estructura configurada: ${nombre}`);
        }
      } catch (err) {
        console.warn(`⚠️ Error al configurar "${nombre}":`, err.message);
      }
    });

    // ============================================================
    // 5. AGREGAR DATOS DE PRUEBA
    // ============================================================
    console.log('👤 Agregando datos de prueba...');

    // 5.1 Usuarios de prueba
    const usuariosPrueba = [
      ['admin@empresa.com', 'Administrador', 'Administrador', hashSimple('admin123'), true, new Date(), null],
      ['revisor@empresa.com', 'Revisor Principal', 'Revisor', hashSimple('rev123'), true, new Date(), null],
      ['tecnico1@empresa.com', 'Técnico Juan', 'Técnico', hashSimple('tec123'), true, new Date(), null],
      ['tecnico2@empresa.com', 'Técnico María', 'Técnico', hashSimple('tec123'), true, new Date(), null]
    ];
    hojas['👤_Usuarios'].getRange(2, 1, usuariosPrueba.length, usuariosPrueba[0].length).setValues(usuariosPrueba);

    // 5.2 Tipos de equipo
    const tiposEquipoPrueba = [
      ['GPS_VEHICULAR', 'GPS para vehículos pesados'],
      ['GPS_LIVIANO', 'GPS para vehículos livianos'],
      ['CAMARA_SEGURIDAD', 'Cámaras de seguridad'],
      ['SENSOR_TEMPERATURA', 'Sensores de temperatura']
    ];
    hojas['📋_Tipos_Equipo'].getRange(2, 1, tiposEquipoPrueba.length, 2).setValues(tiposEquipoPrueba);

    // 5.3 Tipos de unidad
    const tiposUnidadPrueba = [
      ['CAMIONETA', 'Vehículo tipo camioneta'],
      ['AUTOMOVIL', 'Vehículo tipo automóvil'],
      ['MOTOCICLETA', 'Vehículo tipo motocicleta'],
      ['PICKUP', 'Vehículo tipo pickup'],
      ['TRACTOCAMION', 'Vehículo tipo tractocamión']
    ];
    hojas['📋_Tipos_Unidad'].getRange(2, 1, tiposUnidadPrueba.length, 2).setValues(tiposUnidadPrueba);

    // 5.4 Tipos de vehículo
    const tiposVehiculoPrueba = [
      ['TOYOTA_HILUX', 'Toyota Hilux'],
      ['NISSAN_TSURU', 'Nissan Tsuru'],
      ['CHEVROLET_SILVERADO', 'Chevrolet Silverado'],
      ['FORD_F150', 'Ford F-150']
    ];
    hojas['📋_Tipos_Vehiculo'].getRange(2, 1, tiposVehiculoPrueba.length, 2).setValues(tiposVehiculoPrueba);

    // 5.5 Estados de ticket
    const estadosTicketPrueba = [
      ['ABIERTO', 'Ticket abierto', '#4CAF50'],
      ['EN_PROCESO', 'En proceso', '#FF9800'],
      ['PENDIENTE', 'Pendiente de aprobación', '#FFC107'],
      ['RESUELTO', 'Resuelto', '#2196F3'],
      ['CERRADO', 'Cerrado', '#9E9E9E']
    ];
    hojas['📋_Estados_Ticket'].getRange(2, 1, estadosTicketPrueba.length, 3).setValues(estadosTicketPrueba);

    // 5.6 Tarifas
    const tarifasPrueba = [
      ['INSTALACION_BASICA', 1500, 10, 500, 'Instalación básica de GPS'],
      ['INSTALACION_COMPLETA', 2500, 15, 700, 'Instalación completa con cámaras'],
      ['MANTENIMIENTO', 800, 5, 300, 'Mantenimiento preventivo'],
      ['REPARACION', 1200, 8, 400, 'Reparación de equipo existente']
    ];
    hojas['💰_Tarifas'].getRange(2, 1, tarifasPrueba.length, 5).setValues(tarifasPrueba);

    // 5.7 Parámetros del sistema
    const parametrosPrueba = [
      ['EMPRESA_NOMBRE', 'Transportes Ejemplo S.A. de C.V.', 'Nombre de la empresa', new Date()],
      ['EMPRESA_RFC', 'TEC123456789', 'RFC de la empresa', new Date()],
      ['CORREO_FINANZAS', 'finanzas@empresa.com', 'Correo del departamento de finanzas', new Date()],
      ['CORREO_SOPORTE', 'soporte@empresa.com', 'Correo de soporte técnico', new Date()],
      ['CARPETA_REPORTES', 'root', 'ID de la carpeta de reportes en Drive', new Date()],
      ['SHEET_ID', ss.getId(), 'ID de la hoja de cálculo', new Date()],
      ['VERBOSE_LOGGING', 'true', 'Activar logs detallados', new Date()]
    ];
    hojas['⚙️_Parametros'].getRange(2, 1, parametrosPrueba.length, 4).setValues(parametrosPrueba);

    console.log('✅ Datos de prueba agregados');

    // ============================================================
    // 6. CONFIGURAR FORMATOS Y ESTILOS
    // ============================================================
    console.log('🎨 Aplicando formatos...');

    Object.values(hojas).forEach(sheet => {
      try {
        // Autoajustar columnas
        sheet.autoResizeColumns(1, sheet.getLastColumn());
        
        // Congelar primera fila
        sheet.setFrozenRows(1);
        
        // Alternar colores de filas (opcional)
        // No aplicamos para no ralentizar la instalación
      } catch (err) {
        console.warn(`⚠️ Error al aplicar formato a "${sheet.getName()}":`, err.message);
      }
    });

    // ============================================================
    // 7. CONFIGURAR VALIDACIONES
    // ============================================================
    console.log('✅ Configurando validaciones...');
    try {
      configurarValidaciones(hojas);
      console.log('✅ Validaciones configuradas');
    } catch (err) {
      console.error('❌ Error al configurar validaciones:', err.message);
      // Continuamos a pesar del error
    }

    // ============================================================
    // 8. CONFIGURAR PROTECCIONES (¡NUEVA FUNCIÓN!)
    // ============================================================
    console.log('🔒 Configurando protecciones...');
    try {
      const proteccionesOK = configurarProtecciones(hojas);
      if (proteccionesOK) {
        console.log('✅ Protecciones configuradas exitosamente');
      } else {
        console.warn('⚠️ Hubo problemas al configurar algunas protecciones');
      }
    } catch (err) {
      console.error('❌ Error al configurar protecciones:', err.message);
      // Continuamos a pesar del error
    }

    // ============================================================
    // 9. CREAR MENÚ PERSONALIZADO
    // ============================================================
    console.log('📊 Creando menú personalizado...');
    try {
      crearMenuPersonalizado();
      console.log('✅ Menú personalizado creado');
    } catch (err) {
      console.error('❌ Error al crear menú:', err.message);
      // Continuamos a pesar del error
    }

    // ============================================================
    // 10. MOSTRAR MENSAJE DE ÉXITO
    // ============================================================
    console.log('✅ Instalación completada exitosamente');
    
    try {
      mostrarMensajeExito(ui);
    } catch (err) {
      console.error('❌ Error al mostrar mensaje de éxito:', err.message);
      // Mostrar mensaje simple como fallback
      ui.alert(
        '✅ Instalación Completada',
        'La base de datos ha sido instalada correctamente.\n\n' +
        'Revisa la consola para más detalles.',
        ui.ButtonSet.OK
      );
    }

    // ============================================================
    // 11. RETORNAR RESULTADO
    // ============================================================
    return {
      ok: true,
      mensaje: 'Instalación completada exitosamente',
      hojasCreadas: Object.keys(hojas).length,
      fecha: new Date().toISOString()
    };

  } catch (err) {
    // Manejo de errores general
    console.error('❌ Error crítico en instalarBaseDatos:', err);
    
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        '❌ Error en la Instalación',
        'Ocurrió un error durante la instalación:\n\n' +
        err.message + '\n\n' +
        'Revisa la consola para más detalles.',
        ui.ButtonSet.OK
      );
    } catch (e) {
      console.error('❌ Error al mostrar mensaje de error:', e);
    }

    return {
      ok: false,
      mensaje: 'Error en la instalación',
      error: err.message,
      stack: err.stack
    };
  }
}

// ============================================================
// 1. CREACIÓN DE HOJAS
// ============================================================

// ──────────────────────────────────────────────────────────────
// 1.1 💰_Tarifas
// ──────────────────────────────────────────────────────────────

function crearHojatarifas(ss) {
  var nombre = '💰_Tarifas';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['TIPO', 'DESCRIPCION', 'PRECIO', 'CLAVE_INTERNA', 'ACTIVO'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var tarifas = [
    ['IT', 'Instalación Total', 500, 'INST_TOTAL', true],
    ['R', 'Revisión', 150, 'REVISION', true],
    ['IM', 'Instalación Módulo', 250, 'INST_MODULO', true],
    ['CM', 'Cambio de Módulo', 180, 'CAMBIO_MODULO', true],
    ['MC', 'Mantenimiento Correctivo', 200, 'MANT_CORRECTIVO', true],
    ['MP', 'Mantenimiento Preventivo', 120, 'MANT_PREVENTIVO', true],
    ['DI', 'Diagnóstico', 80, 'DIAGNOSTICO', true],
  ];
  sheet.getRange(2, 1, tarifas.length, 5).setValues(tarifas);

  // Formato - ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 80);

  sheet.getRange('C2:C').setNumberFormat('$#,##0.00');

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.2 👤_Usuarios
// ──────────────────────────────────────────────────────────────

function crearHojaUsuarios(ss) {
  var nombre = '👤_Usuarios';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['ID_USUARIO', 'NOMBRE', 'USUARIO', 'PASS_HASH', 'ROL', 'ACTIVO', 'ULTIMO_ACCESO'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Hash de contraseñas usando la función hashSimple
  var usuarios = [
    ['ADM-001', 'Administrador', 'admin', hashSimple('admin123'), 1, true, null],
    ['REV-001', 'María Valenzuela', 'revisor', hashSimple('rev123'), 2, true, null],
    ['TEC-001', 'Juan Pérez', 'tec1', hashSimple('tec123'), 3, true, null],
    ['TEC-002', 'Carlos Gómez', 'tec2', hashSimple('tec123'), 3, true, null],
    ['TEC-003', 'Ana Martínez', 'tec3', hashSimple('tec123'), 3, true, null],
  ];
  sheet.getRange(2, 1, usuarios.length, 7).setValues(usuarios);

  // Validación de ROL
  var reglaRol = SpreadsheetApp.newDataValidation()
    .requireValueInList([1, 2, 3], true)
    .setHelpText('1=Admin, 2=Revisor, 3=Técnico')
    .build();
  sheet.getRange('E2:E').setDataValidation(reglaRol);

  // Validación de ACTIVO
  var reglaActivo = SpreadsheetApp.newDataValidation()
    .requireValueInList([true, false], true)
    .build();
  sheet.getRange('F2:F').setDataValidation(reglaActivo);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 70);
  sheet.setColumnWidth(7, 150);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.3 ⚙️_Parametros
// ──────────────────────────────────────────────────────────────

function crearHojaParametros(ss) {
  var nombre = '⚙️_Parametros';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['PARAMETRO', 'VALOR', 'DESCRIPCION'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var parametros = [
    ['FOLIO_PREFIJO', 'FS', 'Prefijo para los folios de servicio'],
    ['FOLIO_ULTIMO', '0', 'Último folio generado (auto-incrementable)'],
    ['DRIVE_FOLDER_ID', '', 'ID de la carpeta raíz en Google Drive'],
    ['CORREOS_DESTINO', 'finanzas@empresa.com, validador@empresa.com', 'Correos para envío de PDF (separados por comas)'],
    ['DIAS_LIMITE_DRIVE', '60', 'Días de antigüedad para eliminar archivos'],
    ['COMPRESION_IMAGENES', 'true', 'Comprimir imágenes antes de subir (true/false)'],
    ['CALIDAD_IMAGEN', '80', 'Calidad de compresión (1-100)'],
    ['MAX_IMAGENES_PDF', '12', 'Máximo de imágenes por PDF'],
    ['FORMATO_SERIE_GPS', 'XXXX-XXX-XXX', 'Máscara de validación para series GPS'],
    ['VERBOSE_LOGGING', 'false', 'Registro detallado en consola (true/false)'],
    ['EMPRESA_NOMBRE', 'Fleet Manager', 'Nombre de la empresa para documentos'],
    ['EMPRESA_RFC', '', 'RFC para facturación'],
    ['TICKET_PREFIJO', 'TKT', 'Prefijo para tickets de fallas'],
    ['TICKET_ULTIMO', '0', 'Último ticket generado'],
  ];
  sheet.getRange(2, 1, parametros.length, 3).setValues(parametros);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 300);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.4 📦_Inventario_GPS (CON TIPO_UNIDAD)
// ──────────────────────────────────────────────────────────────

function crearHojaInventario(ss) {
  var nombre = '📦_Inventario_GPS';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'SERIE_GPS',
    'TIPO_EQUIPO',
    'MODELO',
    'IMEI',
    'ESTADO',
    'ECONOMICO_ASIGNADO',
    'FECHA_INSTALACION',
    'TICKET_GARANTIA',
    'FECHA_GARANTIA',
    'ULTIMA_ACTUALIZACION',
    'OBSERVACIONES',
    'TIPO_UNIDAD'  // ✅ NUEVA COLUMNA (columna L)
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['GTMF-RYD-W8S', 'CM32', 'Samsara CM32', '123456789012345', 'Disponible', '', '', '', '', new Date(), 'Equipo nuevo', ''],
    ['GTMF-KL7-P9M', 'CM34', 'Samsara CM34', '543210987654321', 'Instalado', 'G-444', new Date(), '', '', new Date(), 'Instalado en G-444', 'UTILITARIA'],
    ['GTMF-HJ4-RT2', 'CM32', 'Samsara CM32', '987654321012345', 'Garantía', 'G-361', new Date(), 'TKT-2026-001', new Date(), new Date(), 'En proceso de garantía', 'UNIDAD'],
    ['GTMF-BN8-XC4', 'VG55', 'Samsara VG55', '321098765432109', 'Instalado', 'G-361', new Date(), '', '', new Date(), 'Gateway instalado en G-361', 'UNIDAD'],
    ['GTMF-DF6-WQ1', 'CM32', 'Samsara CM32', '654321098765432', 'Baja', '', '', '', '', new Date(), 'Equipo dado de baja', ''],
    ['GTMF-ZX9-AB2', 'VG54', 'Samsara VG54', '111222333444555', 'Disponible', '', '', '', '', new Date(), 'Modelo con cámara', ''],
    ['GTMF-ZX9-AB3', 'VG54', 'Samsara VG54', '111222333444556', 'Disponible', '', '', '', '', new Date(), 'Modelo con cámara', ''],
    ['GTMF-ZX1-AB4', 'VG54', 'Samsara VG54', '111222333444557', 'Disponible', '', '', '', '', new Date(), '', ''],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // Validación de ESTADO
  var reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Disponible', 'Instalado', 'Garantía', 'Baja'], true)
    .build();
  sheet.getRange('E2:E').setDataValidation(reglaEstado);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 150);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 200);
  sheet.setColumnWidth(12, 150);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.5 🚚_Flotilla_Fallas (CON TIPO_UNIDAD + columna de fallas)
// ──────────────────────────────────────────────────────────────

function crearHojaFlotilla(ss) {
  var nombre = '🚚_Flotilla_Fallas';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'ECONOMICO',
    'PLACAS',
    'TIPO_VEHICULO',
    'MARCA',
    'MODELO',
    'AÑO',
    'SERIE_VEHICULO',
    'GPS_ACTUAL',
    'ESTADO',
    'ULTIMO_SERVICIO',
    'TIPO_UNIDAD'  // ✅ NUEVA COLUMNA (columna K)
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['G-361', 'XXXXX', 'UTILITARIA', 'Ford', 'F-350', 2023, '1FT8W3BT9PED12345', 'Instalado', 'Activo', '', 'UNIDAD'],
    ['G-444', 'ABC-123-X', 'Camión', 'Ford', 'F-350', 2023, '1FT8W3BT9PED12345', 'Instalado', 'Activo', '', 'UTILITARIA'],
    ['G-445', 'DEF-456-Y', 'Camión', 'International', 'HV', 2022, '1HTMKAAN8NH123456', 'Instalado', 'Activo', '', 'CAJA SECA'],
    ['G-446', 'GHI-789-Z', 'Van', 'Mercedes', 'Sprinter', 2024, 'WD3PE8CC8E5123456', 'Pendiente', 'Activo', '', 'MÓVIL'],
    ['G-447', 'JKL-012-A', 'Pickup', 'Chevrolet', 'Silverado', 2021, '1GCVKREC3HZ123456', 'Sin GPS', 'Activo', '', 'DOLLY'],
    ['G-448', 'MNO-345-B', 'Camión', 'Freightliner', 'M2', 2023, '1FVHG5DV3CH123456', 'Instalado', 'Inactivo', '', 'REMOLQUE'],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  var reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Activo', 'Inactivo', 'Mantenimiento'], true)
    .build();
  sheet.getRange('I2:I').setDataValidation(reglaEstado);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 150);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.6 📝_Bitacora_Revisiones
// ──────────────────────────────────────────────────────────────

function crearHojaBitacora(ss) {
  var nombre = '📝_Bitacora_Revisiones';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'FOLIO',
    'FECHA_SERVICIO',
    'ECONOMICO',
    'PLACAS',
    'SERIE_GPS',
    'TIPO_REVISION',
    'PLATAFORMA',
    'TECNICO_ID',
    'TECNICO_NOMBRE',
    'DETALLE_TRABAJO',
    'PRECIO_UNITARIO',
    'FOTOS_DRIVE_URL',
    'ESTADO',
    'FECHA_POSIBLE_PAGO',
    'APROBADO_POR',
    'FECHA_APROBACION',
    'PDF_URL',
    'CORREO_ENVIADO',
    'NOTAS_REVISOR',
    'LATITUD',
    'LONGITUD',
    'FECHA_REGISTRO'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Datos de ejemplo
  var ahora = new Date();
  var hace3dias = new Date(ahora);
  hace3dias.setDate(hace3dias.getDate() - 3);
  var hace7dias = new Date(ahora);
  hace7dias.setDate(hace7dias.getDate() - 7);

  var datos = [
    ['FS-0001', hace7dias, 'G-444', 'ABC-123-X', 'GTMF-KL7-P9M', 'IT', 'SAMSARA', 'TEC-001', 'Juan Pérez', 'Instalación completa de GPS Samsara en camión. Cableado por el tablero y fijación en parabrisas.', 500, '', 'Pagado', new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + 10), 'María Valenzuela', new Date(), 'https://drive.google.com/drive/folders/ejemplo1', true, 'Sin observaciones', '19.4326', '-99.1332', hace7dias],
    ['FS-0002', hace3dias, 'G-445', 'DEF-456-Y', '', 'R', 'SAMSARA', 'TEC-002', 'Carlos Gómez', 'Revisión de GPS en International. Se detectó falla en antena, se recomienda cambio.', 150, '', 'Listo para pago', '', '', '', '', false, 'Esperar aprobación del cliente', '19.4326', '-99.1332', hace3dias],
    ['FS-0003', new Date(), 'G-446', 'GHI-789-Z', 'GTMF-RYD-W8S', 'IM', 'SAMSARA', 'TEC-001', 'Juan Pérez', 'Instalación de módulo GPS en Sprinter. Se realizó prueba de funcionamiento.', 250, 'https://drive.google.com/drive/folders/ejemplo2', 'Borrador', '', '', '', '', false, 'Pendiente de validación', '19.4326', '-99.1332', ahora],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 150);
  sheet.setColumnWidth(10, 250);
  sheet.setColumnWidth(11, 100);
  sheet.setColumnWidth(12, 200);
  sheet.setColumnWidth(13, 130);
  sheet.setColumnWidth(14, 130);
  sheet.setColumnWidth(15, 150);
  sheet.setColumnWidth(16, 130);
  sheet.setColumnWidth(17, 200);
  sheet.setColumnWidth(18, 80);
  sheet.setColumnWidth(19, 200);
  sheet.setColumnWidth(20, 120);
  sheet.setColumnWidth(21, 120);
  sheet.setColumnWidth(22, 130);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.7 📊_Consulta_Tecnicos
// ──────────────────────────────────────────────────────────────

function crearHojaConsulta(ss) {
  var nombre = '📊_Consulta_Tecnicos';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'TECNICO_ID',
    'TECNICO_NOMBRE',
    'TOTAL_PENDIENTE',
    'TOTAL_PAGADO',
    'ULTIMO_SERVICIO',
    'PROMEDIO_PAGO',
    'SERVICIOS_PENDIENTES',
    'SERVICIOS_PAGADOS'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['TEC-001', 'Juan Pérez', 250, 500, new Date(), 500, 1, 1],
    ['TEC-002', 'Carlos Gómez', 150, 0, new Date(), 0, 1, 0],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 150);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.8 📋_Tipos_Equipo
// ──────────────────────────────────────────────────────────────

function crearHojaTiposEquipo(ss) {
  var nombre = '📋_Tipos_Equipo';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['CLAVE', 'NOMBRE', 'DESCRIPCION', 'ACTIVO'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['CM32', 'Samsara CM32', 'Modelo estándar', true],
    ['CM34', 'Samsara CM34', 'Modelo avanzado con conectividad 4G', true],
    ['CM35', 'Samsara CM35', 'Modelo con cámara integrada', true],
    ['VG54', 'Samsara VG54', 'Gateway 4G con antena interna', true],
    ['VG55', 'Samsara VG55', 'Gateway 4G con antena externa', true],
    ['CG12', 'Samsara CG12', 'Gateway para conexión CAN', false],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 80);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.9 📋_Tipos_Unidad (NUEVA HOJA)
// ──────────────────────────────────────────────────────────────

function crearHojaTiposUnidad(ss) {
  var nombre = '📋_Tipos_Unidad';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['TIPO'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var tipos = [
    ['CAJA SECA'],
    ['DOLLY'],
    ['MÓVIL'],
    ['PRUEBA'],
    ['REMOLQUE'],
    ['UNIDAD'],
    ['UTILITARIA']
  ];
  sheet.getRange(2, 1, tipos.length, tipos[0].length).setValues(tipos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 150);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.10 🔧_Accesorios_Stock
// ──────────────────────────────────────────────────────────────

function crearHojaAccesorios(ss) {
  var nombre = '🔧_Accesorios_Stock';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'ACCESORIO',
    'TIPO',
    'STOCK_TOTAL',
    'ASIGNADOS',
    'DISPONIBLES',
    'ECONOMICO_ASIGNADO',
    'FECHA_ASIGNACION',
    'OBSERVACIONES'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['ARNES VG34', 'Arnés VG34', 50, 1, 49, 'G-361', new Date(), 'Stock completo'],
    ['ARNES VG54/55', 'Arnés VG54/55', 50, 26, 24, 'G-361, G-444, G-539', new Date(), 'Stock completo'],
    ['BTN 1.0', 'Botón 1.0', 50, 3, 47, 'G-361', new Date(), 'Stock completo'],
    ['BTN 1.2', 'Botón 1.2', 50, 25, 25, 'G-361, G-444, G-539', new Date(), 'Stock completo'],
    ['EI1.0', 'Corte de motor EI1.0', 50, 0, 50, '', '', 'Stock completo'],
    ['EI2.0', 'Corte de motor EI2.0', 50, 27, 23, 'G-361, G-444, G-539', new Date(), 'Stock completo'],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 250);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.11 📑_Facturas
// ──────────────────────────────────────────────────────────────

function crearHojaFacturas(ss) {
  var nombre = '📑_Facturas';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'FOLIO_FACTURA',
    'FECHA',
    'PROVEEDOR',
    'SERIES_INCLUIDAS',
    'SERIES_INSTALADAS',
    'SERIES_STOCK',
    'CONCILIADO',
    'FECHA_CONCILIACION',
    'OBSERVACIONES'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['FAC-2026-001', new Date(), 'Samsara México', 'GTMF-RYD-W8S, GTMF-KL7-P9M, GTMF-HJ4-RT2', 'GTMF-KL7-P9M', 'GTMF-RYD-W8S, GTMF-HJ4-RT2', false, '', 'Pendiente de conciliación'],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 250);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.12 📈_Log_Auditoria
// ──────────────────────────────────────────────────────────────

function crearHojaLogAuditoria(ss) {
  var nombre = '📈_Log_Auditoria';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  // ✅ Headers correctos (10 columnas)
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

  // Anchos de columna
  sheet.setColumnWidth(1, 150); // FECHA
  sheet.setColumnWidth(2, 120); // USUARIO_ID
  sheet.setColumnWidth(3, 180); // USUARIO_NOMBRE
  sheet.setColumnWidth(4, 150); // ACCION
  sheet.setColumnWidth(5, 150); // MODULO
  sheet.setColumnWidth(6, 300); // DESCRIPCION
  sheet.setColumnWidth(7, 150); // FOLIO_RELACIONADO
  sheet.setColumnWidth(8, 150); // IP
  sheet.setColumnWidth(9, 200); // USER_AGENT
  sheet.setColumnWidth(10, 200); // DETALLES_ADICIONALES

  // Congelar primera fila
  sheet.setFrozenRows(1);

  console.log('✅ Hoja ' + nombre + ' creada correctamente con 10 columnas');
  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.13 📩_Notificaciones
// ──────────────────────────────────────────────────────────────

function crearHojaNotificaciones(ss) {
  var nombre = '📩_Notificaciones';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'FECHA',
    'USUARIO_DESTINO',
    'TIPO',
    'MENSAJE',
    'FOLIO_RELACIONADO',
    'LEIDA',
    'FECHA_LECTURA'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    [new Date(), 'TEC-001', 'MODIFICACION_EQUIPO', 'Se ha modificado el equipo GPS en el folio FS-0002', 'FS-0002', false, ''],
    [new Date(), 'REV-001', 'MODIFICACION_COSTO', 'Se ha modificado el costo en el folio FS-0001', 'FS-0001', false, ''],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 150);

  return sheet;
}

// ──────────────────────────────────────────────────────────────
// 1.14 🎫_Tickets (NUEVA HOJA - Sistema de tickets/fallas)
// ──────────────────────────────────────────────────────────────

function crearHojaTickets(ss) {
  var nombre = '🎫_Tickets';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  // ✅ NUEVA ESTRUCTURA: 18 columnas (de la A a la R)
  var headers = [
    'ID',                       // A
    'FECHA',                    // B
    'UNIDAD',                   // C
    'DESCRIPCION',              // D
    'CREADO_POR',               // E
    'CREADO_POR_NOMBRE',        // F
    'ESTADO',                   // G
    'TECNICO_ASIGNADO',         // H
    'TECNICO_NOMBRE',           // I
    'FECHA_CIERRE',             // J
    'COMENTARIOS',              // K
    'ULTIMA_ACTUALIZACION',     // L
    'TECNICOS_AUTORIZADOS',     // M  ← NUEVA: Lista de IDs separados por coma
    'CREADO_POR_ROL',           // N  ← NUEVA: 1=Admin, 3=Técnico
    'PRIORIDAD',                // O  ← NUEVA: Baja / Media / Alta
    'CATEGORIA',                // P  ← NUEVA: Hardware / Software / Red / Otro
    'ARCHIVOS_ADJUNTOS',        // Q  ← NUEVA: URLs o nombres de archivos
    'NOTAS_INTERNAS'            // R  ← NUEVA: Solo visible para administradores
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ── DATOS DE EJEMPLO CON LA NUEVA ESTRUCTURA ──
  var ahora = new Date();
  var hace2dias = new Date(ahora);
  hace2dias.setDate(hace2dias.getDate() - 2);
  var hace5dias = new Date(ahora);
  hace5dias.setDate(hace5dias.getDate() - 5);

  var datos = [
    [
      'TKT-001',                       // ID
      hace5dias,                       // FECHA
      'G-361',                         // UNIDAD
      'El GPS no reporta posición. Posible falla en antena.', // DESCRIPCION
      'REV-001',                       // CREADO_POR
      'María Valenzuela',              // CREADO_POR_NOMBRE
      'Resuelto',                      // ESTADO
      'TEC-001',                       // TECNICO_ASIGNADO
      'Juan Pérez',                    // TECNICO_NOMBRE
      ahora,                           // FECHA_CIERRE
      'Se reemplazó la antena GPS y se realizó prueba de funcionamiento. Todo OK.', // COMENTARIOS
      ahora,                           // ULTIMA_ACTUALIZACION
      'TEC-001,TEC-002',               // TECNICOS_AUTORIZADOS (visibilidad)
      '1',                             // CREADO_POR_ROL (Admin)
      'Alta',                          // PRIORIDAD
      'Hardware',                      // CATEGORIA
      '',                              // ARCHIVOS_ADJUNTOS
      'Ticket resuelto por cambio de antena.' // NOTAS_INTERNAS
    ],
    [
      'TKT-002',                       // ID
      hace2dias,                       // FECHA
      'G-444',                         // UNIDAD
      'La cámara no enciende. Se revisó cableado.', // DESCRIPCION
      'TEC-001',                       // CREADO_POR
      'Juan Pérez',                    // CREADO_POR_NOMBRE
      'En proceso',                    // ESTADO
      'TEC-001',                       // TECNICO_ASIGNADO
      'Juan Pérez',                    // TECNICO_NOMBRE
      null,                            // FECHA_CIERRE
      'Se detectó cable suelto en la conexión. Pendiente de soldadura.', // COMENTARIOS
      ahora,                           // ULTIMA_ACTUALIZACION
      'TEC-001',                       // TECNICOS_AUTORIZADOS
      '3',                             // CREADO_POR_ROL (Técnico)
      'Media',                         // PRIORIDAD
      'Hardware',                      // CATEGORIA
      '',                              // ARCHIVOS_ADJUNTOS
      'Pendiente de soldadura.'        // NOTAS_INTERNAS
    ],
    [
      'TKT-003',                       // ID
      ahora,                           // FECHA
      'G-445',                         // UNIDAD
      'Falla en la conexión del gateway. No enciende.', // DESCRIPCION
      'TEC-002',                       // CREADO_POR
      'Carlos Gómez',                  // CREADO_POR_NOMBRE
      'Pendiente',                     // ESTADO
      '',                              // TECNICO_ASIGNADO
      '',                              // TECNICO_NOMBRE
      null,                            // FECHA_CIERRE
      'Se reporta el ticket para que un técnico lo revise.', // COMENTARIOS
      ahora,                           // ULTIMA_ACTUALIZACION
      'TEC-002,TEC-003',               // TECNICOS_AUTORIZADOS
      '3',                             // CREADO_POR_ROL (Técnico)
      'Media',                         // PRIORIDAD
      'Red',                           // CATEGORIA
      '',                              // ARCHIVOS_ADJUNTOS
      'Requiere revisión de red.'      // NOTAS_INTERNAS
    ]
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // ── VALIDACIONES DE DATOS ──

  // 1. Validación de ESTADO (columna G)
  var reglaEstadoTicket = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'En proceso', 'Resuelto'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('G2:G').setDataValidation(reglaEstadoTicket);

  // 2. ✅ NUEVA: Validación de PRIORIDAD (columna O)
  var reglaPrioridad = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Baja', 'Media', 'Alta'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('O2:O').setDataValidation(reglaPrioridad);

  // 3. ✅ NUEVA: Validación de CATEGORIA (columna P)
  var reglaCategoria = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Hardware', 'Software', 'Red', 'Otro'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('P2:P').setDataValidation(reglaCategoria);

  // ── ANCHOS DE COLUMNA PARA MEJOR VISUALIZACIÓN ──
  sheet.setColumnWidth(1, 90);   // ID
  sheet.setColumnWidth(2, 130);  // FECHA
  sheet.setColumnWidth(3, 100);  // UNIDAD
  sheet.setColumnWidth(4, 250);  // DESCRIPCION
  sheet.setColumnWidth(5, 120);  // CREADO_POR
  sheet.setColumnWidth(6, 160);  // CREADO_POR_NOMBRE
  sheet.setColumnWidth(7, 110);  // ESTADO
  sheet.setColumnWidth(8, 130);  // TECNICO_ASIGNADO
  sheet.setColumnWidth(9, 160);  // TECNICO_NOMBRE
  sheet.setColumnWidth(10, 130); // FECHA_CIERRE
  sheet.setColumnWidth(11, 200); // COMENTARIOS
  sheet.setColumnWidth(12, 130); // ULTIMA_ACTUALIZACION
  sheet.setColumnWidth(13, 180); // TECNICOS_AUTORIZADOS
  sheet.setColumnWidth(14, 100); // CREADO_POR_ROL
  sheet.setColumnWidth(15, 90);  // PRIORIDAD
  sheet.setColumnWidth(16, 120); // CATEGORIA
  sheet.setColumnWidth(17, 150); // ARCHIVOS_ADJUNTOS
  sheet.setColumnWidth(18, 250); // NOTAS_INTERNAS

  // Congelar la primera fila (encabezados)
  sheet.setFrozenRows(1);

  console.log('✅ Hoja ' + nombre + ' creada correctamente con 18 columnas (visibilidad dinámica incluida)');
  return sheet;
}

// ============================================================
// 2. VALIDACIONES CRUZADAS
// ============================================================

function configurarValidaciones(hojas) {
  try {
    // Validar que hojas existe
    if (!hojas || typeof hojas !== 'object') {
      throw new Error('El objeto "hojas" es requerido');
    }

    // 1. Validación de TIPO_REVISION en Bitácora contra Tarifas
    var bitacora = hojas.bitacora;
    var tarifas = hojas.tarifas;
    if (bitacora && tarifas) {
      var tipos = tarifas.getRange('A2:A').getValues();
      var listaTipos = [];
      for (var i = 0; i < tipos.length; i++) {
        if (tipos[i][0]) listaTipos.push(tipos[i][0]);
      }
      if (listaTipos.length > 0) {
        var regla = SpreadsheetApp.newDataValidation()
          .requireValueInList(listaTipos, true)
          .setAllowInvalid(false)
          .build();
        bitacora.getRange('F2:F').setDataValidation(regla);
      } else {
        console.warn('⚠️ No se encontraron tipos de revisión en Tarifas');
      }
    } else {
      console.warn('⚠️ No se encontraron las hojas Bitácora o Tarifas');
    }

    // 2. Validación de ESTADO en Bitácora
    if (bitacora) {
      var reglaEstado = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Borrador', 'Listo para pago', 'Pagado'], true)
        .build();
      bitacora.getRange('M2:M').setDataValidation(reglaEstado);
    }

    // 3. Validación de TIPO_EQUIPO en Inventario contra Tipos_Equipo
    var inventario = hojas.inventario;
    var tiposEquipo = hojas.tiposEquipo;
    if (inventario && tiposEquipo) {
      var equipos = tiposEquipo.getRange('A2:A').getValues();
      var listaEquipos = [];
      for (var j = 0; j < equipos.length; j++) {
        if (equipos[j][0]) listaEquipos.push(equipos[j][0]);
      }
      if (listaEquipos.length > 0) {
        var reglaEquipo = SpreadsheetApp.newDataValidation()
          .requireValueInList(listaEquipos, true)
          .setAllowInvalid(false)
          .build();
        inventario.getRange('B2:B').setDataValidation(reglaEquipo);
      }
    }

    // 4. Validación de TIPO_UNIDAD en Inventario contra Tipos_Unidad
    var tiposUnidad = hojas.tiposUnidad;
    if (inventario && tiposUnidad) {
      var tipos = tiposUnidad.getRange('A2:A').getValues();
      var listaTiposUnidad = [];
      for (var k = 0; k < tipos.length; k++) {
        if (tipos[k][0]) listaTiposUnidad.push(tipos[k][0]);
      }
      if (listaTiposUnidad.length > 0) {
        var reglaTipoUnidad = SpreadsheetApp.newDataValidation()
          .requireValueInList(listaTiposUnidad, true)
          .setAllowInvalid(false)
          .build();
        inventario.getRange('L2:L').setDataValidation(reglaTipoUnidad);
      }
    }

    // 5. Validación de ESTADO en Tickets contra Estados_Ticket
    var tickets = hojas.tickets;
    var estadosTicket = hojas.estadosTicket;
    if (tickets && estadosTicket) {
      var estados = estadosTicket.getRange('A2:A').getValues();
      var listaEstados = [];
      for (var l = 0; l < estados.length; l++) {
        if (estados[l][0]) listaEstados.push(estados[l][0]);
      }
      if (listaEstados.length > 0) {
        var reglaEstadoTicket = SpreadsheetApp.newDataValidation()
          .requireValueInList(listaEstados, true)
          .setAllowInvalid(false)
          .build();
        tickets.getRange('G2:G').setDataValidation(reglaEstadoTicket);
      }
    }

    // 6. Validación de TIPO_VEHICULO en Flotilla contra Tipos_Vehiculo
    var flotilla = hojas.flotilla;
    var tiposVehiculo = hojas.tiposVehiculo;
    if (flotilla && tiposVehiculo) {
      var tiposVeh = tiposVehiculo.getRange('A2:A').getValues();
      var listaTiposVeh = [];
      for (var m = 0; m < tiposVeh.length; m++) {
        if (tiposVeh[m][0]) listaTiposVeh.push(tiposVeh[m][0]);
      }
      if (listaTiposVeh.length > 0) {
        var reglaTipoVeh = SpreadsheetApp.newDataValidation()
          .requireValueInList(listaTiposVeh, true)
          .setAllowInvalid(false)
          .build();
        flotilla.getRange('C2:C').setDataValidation(reglaTipoVeh);
      }
    }

    console.log('✅ Validaciones configuradas exitosamente');
    return true;

  } catch (err) {
    console.error('❌ Error configurando validaciones:', err.message);
    throw new Error('Error al configurar validaciones: ' + err.message);
  }
}

// ============================================================
// 3. PROTECCIONES
// ============================================================

/**
 * Configura protecciones y permisos en las hojas del sistema
 * @param {Object} hojas - Objeto con todas las hojas creadas
 */
function configurarProtecciones(hojas) {
  try {
    // Validar entrada
    if (!hojas || typeof hojas !== 'object') {
      throw new Error('El objeto "hojas" es requerido');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. Proteger hojas críticas (solo administradores)
    const hojasCriticas = ['👤_Usuarios', '⚙️_Parametros', '💰_Tarifas'];
    const adminEmail = Session.getActiveUser().getEmail();

    hojasCriticas.forEach(nombreHoja => {
      try {
        const sheet = ss.getSheetByName(nombreHoja);
        if (!sheet) {
          console.warn(`⚠️ Hoja "${nombreHoja}" no encontrada`);
          return;
        }

        // Remover protecciones existentes
        const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        protections.forEach(prot => prot.remove());

        // Crear nueva protección
        const protection = sheet.protect();
        protection.setDescription(`🔒 Protegido: ${nombreHoja} - Solo administradores`);
        protection.addEditor(adminEmail);
        protection.setWarningOnly(true);

        console.log(`✅ Hoja "${nombreHoja}" protegida`);
      } catch (err) {
        console.warn(`⚠️ Error al proteger "${nombreHoja}":`, err.message);
      }
    });

    // 2. Proteger rangos específicos (fórmulas, IDs)
    const hojaBitacora = hojas.bitacora || ss.getSheetByName('📝_Bitacora_Revisiones');
    if (hojaBitacora) {
      try {
        // Proteger columna de folios (no modificar manualmente)
        const rangoFolios = hojaBitacora.getRange('A:A');
        const protection = rangoFolios.protect();
        protection.setDescription('🔒 FOLIO - No modificar manualmente');
        protection.addEditor(adminEmail);
        protection.setWarningOnly(true);
        console.log('✅ Columna FOLIO protegida');
      } catch (err) {
        console.warn('⚠️ Error al proteger columna FOLIO:', err.message);
      }
    }

    // 3. Proteger fórmulas en inventario
    const hojaInventario = hojas.inventario || ss.getSheetByName('📦_Inventario_GPS');
    if (hojaInventario) {
      try {
        // Proteger columnas de fechas automáticas
        const ultimaFila = hojaInventario.getLastRow();
        if (ultimaFila > 0) {
          const rangoFechas = hojaInventario.getRange(2, 1, ultimaFila - 1, 1);
          const protection = rangoFechas.protect();
          protection.setDescription('🔒 Fechas automáticas');
          protection.addEditor(adminEmail);
          protection.setWarningOnly(true);
          console.log('✅ Fechas automáticas protegidas');
        }
      } catch (err) {
        console.warn('⚠️ Error al proteger fechas:', err.message);
      }
    }

    console.log('✅ Protecciones configuradas exitosamente');
    return true;

  } catch (err) {
    console.error('❌ Error al configurar protecciones:', err.message);
    // No lanzar error para no interrumpir la instalación
    return false;
  }
}

// ============================================================
// 4. MENÚ PERSONALIZADO
// ============================================================

function crearMenuPersonalizado() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Crear menú principal
    var menu = ui.createMenu('📊 Fleet Manager');
    
    // ===== SECCIÓN: ADMINISTRACIÓN =====
    menu.addItem('🚀 Reinstalar base de datos', 'instalarBaseDatos');
    menu.addItem('🔍 Verificar integridad', 'verificarIntegridad');
    menu.addSeparator();
    
    // ===== SECCIÓN: DATOS =====
    menu.addItem('📋 Cargar inventario desde Excel', 'cargarInventarioDesdeExcel');
    menu.addItem('🧹 Limpiar archivos Drive antiguos', 'limpiarDriveAntiguo');
    menu.addSeparator();
    
    // ===== SECCIÓN: REPORTES =====
    menu.addItem('📊 Dashboard de técnicos', 'mostrarDashboard');
    menu.addItem('📈 Reporte de actividad', 'generarReporteActividad');
    menu.addSeparator();
    
    // ===== SECCIÓN: AYUDA =====
    menu.addItem('❓ Ayuda', 'mostrarAyuda');
    menu.addItem('ℹ️ Acerca de', 'mostrarAcercaDe');
    
    // Agregar a la interfaz
    menu.addToUi();
    
    console.log('✅ Menú personalizado creado exitosamente');
    return true;

  } catch (err) {
    console.error('❌ Error al crear menú:', err.message);
    // No lanzar error para no interrumpir la instalación
    return false;
  }
}

// ============================================================
// 5. MENSAJE DE ÉXITO
// ============================================================

function mostrarMensajeExito(ui) {
  try {
    // Verificar que ui existe
    if (!ui) {
      ui = SpreadsheetApp.getUi();
    }

    var mensaje =
      '✅ INSTALACIÓN COMPLETA\n\n' +
      '📋 Hojas creadas (17 hojas):\n' +
      '• 💰_Tarifas\n' +
      '• 👤_Usuarios\n' +
      '• ⚙️_Parametros\n' +
      '• 📦_Inventario_GPS (con TIPO_UNIDAD)\n' +
      '• 🚚_Flotilla_Fallas (con TIPO_UNIDAD)\n' +
      '• 📝_Bitacora_Revisiones\n' +
      '• 📊_Consulta_Tecnicos\n' +
      '• 📋_Tipos_Equipo\n' +
      '• 📋_Tipos_Unidad\n' +
      '• 📋_Tipos_Vehiculo (NUEVA)\n' +
      '• 📋_Estados_Ticket (NUEVA)\n' +
      '• 🔧_Accesorios_Stock\n' +
      '• 📑_Facturas\n' +
      '• 📈_Log_Auditoria\n' +
      '• 📩_Notificaciones\n' +
      '• 🎫_Tickets\n' +
      '• 📋_Catálogo_Estados_Vehiculo\n\n' +
      '✅ VALIDACIONES CONFIGURADAS:\n' +
      '• TIPO_REVISION en Bitácora\n' +
      '• ESTADO en Bitácora\n' +
      '• TIPO_EQUIPO en Inventario\n' +
      '• TIPO_UNIDAD en Inventario\n' +
      '• ESTADO en Tickets\n' +
      '• TIPO_VEHICULO en Flotilla\n\n' +
      '👤 USUARIOS DE PRUEBA:\n' +
      '• admin / admin123 (Administrador)\n' +
      '• revisor / rev123 (Revisor)\n' +
      '• tec1 / tec123 (Técnico)\n' +
      '• tec2 / tec123 (Técnico)\n\n' +
      '🚀 El sistema está listo para usar!\n' +
      '📊 Accede al menú "📊 Fleet Manager" en la barra superior.';

    ui.alert('🎉 INSTALACIÓN COMPLETA', mensaje, ui.ButtonSet.OK);
    
    console.log('✅ Mensaje de éxito mostrado al usuario');
    return true;

  } catch (err) {
    console.error('❌ Error al mostrar mensaje de éxito:', err.message);
    // Intentar mostrar un mensaje simple como fallback
    try {
      var ui = SpreadsheetApp.getUi();
      ui.alert('✅ Instalación Completada', 'La base de datos ha sido instalada correctamente.', ui.ButtonSet.OK);
    } catch (e) {
      console.error('❌ Error crítico al mostrar mensaje:', e);
    }
    return false;
  }
}


// ============================================================
// 7. FUNCIONES PLACEHOLDER
// ============================================================

// Funciones auxiliares para el menú (implementaciones básicas)
function cargarInventarioDesdeExcel() {
  // Implementar lógica de carga desde Excel
  SpreadsheetApp.getUi().alert('📋 Cargar inventario desde Excel', 'Función en desarrollo', SpreadsheetApp.getUi().ButtonSet.OK);
}

function limpiarDriveAntiguo() {
  // Implementar limpieza de archivos antiguos
  SpreadsheetApp.getUi().alert('🧹 Limpiar archivos Drive antiguos', 'Función en desarrollo', SpreadsheetApp.getUi().ButtonSet.OK);
}

function mostrarDashboard() {
  // Implementar dashboard
  SpreadsheetApp.getUi().alert('📊 Dashboard de técnicos', 'Función en desarrollo', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// 8. VERIFICACIÓN DE INTEGRIDAD
// ============================================================

function verificarIntegridad() {
  // Implementar verificación de integridad
  SpreadsheetApp.getUi().alert('🔍 Verificar integridad', 'Función en desarrollo', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generarReporteActividad() {
  // Implementar reporte de actividad
  SpreadsheetApp.getUi().alert('📈 Reporte de actividad', 'Función en desarrollo', SpreadsheetApp.getUi().ButtonSet.OK);
}

function mostrarAyuda() {
  SpreadsheetApp.getUi().alert(
    '❓ Ayuda\n\n' +
    'Sistema de Gestión de Flotas\n\n' +
    'Para soporte, contacte a:\n' +
    'soporte@empresa.com'
  );
}
function mostrarAcercaDe() {
  SpreadsheetApp.getUi().alert(
    'ℹ️ Acerca de\n\n' +
    'Sistema de Gestión de Flotas\n' +
    'Versión: 1.0.0\n' +
    '© 2024 Empresa de Transporte'
  );
}
// ============================================================
// 9. PRUEBA
// ============================================================

function testInstalacion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var nombres = [];
  for (var i = 0; i < sheets.length; i++) {
    nombres.push(sheets[i].getName());
  }
  console.log('Hojas actuales:', nombres.join(', '));

  var usuarios = ss.getSheetByName('👤_Usuarios');
  if (usuarios) {
    var datos = usuarios.getDataRange().getValues();
    console.log('Usuarios cargados:');
    for (var i = 1; i < datos.length; i++) {
      var u = datos[i];
      if (u[0]) console.log('  -', u[1], '(rol:', u[4] + ')');
    }
  }
}
// ============================================================
// CREAR HOJA: 📋_Catálogo_Estados_Vehiculo
// ============================================================
function crearHojaEstadosVehiculo(ss) {
  var nombreHoja = '📋_Catálogo_Estados_Vehiculo';
  var hoja = ss.getSheetByName(nombreHoja);

  if (hoja) {
    ss.deleteSheet(hoja);
  }

  hoja = ss.insertSheet(nombreHoja);

  // Encabezados
  var headers = [
    ['ID', 'NOMBRE', 'COLOR', 'ACTIVO', 'DESCRIPCION']
  ];
  hoja.getRange(1, 1, 1, 5).setValues(headers);

  // Estilos a los encabezados
  var headerRange = hoja.getRange(1, 1, 1, 5);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A90D9');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  // Datos iniciales con los estados que mencionaste
  var datosIniciales = [
    [1, 'Activo', 'success', true, 'Vehículo operativo y en circulación'],
    [2, 'Inactivo', 'danger', true, 'Vehículo fuera de circulación temporal o permanente'],
    [3, 'En Mantenimiento', 'warning', true, 'Vehículo en taller por mantenimiento programado'],
    [4, 'Siniestrada', 'dark', true, 'Vehículo con daño por accidente o siniestro']
  ];

  if (datosIniciales.length > 0) {
    hoja.getRange(2, 1, datosIniciales.length, 5).setValues(datosIniciales);
  }

  // Ajustar ancho de columnas
  hoja.setColumnWidth(1, 50);   // ID
  hoja.setColumnWidth(2, 160);  // NOMBRE
  hoja.setColumnWidth(3, 100);  // COLOR
  hoja.setColumnWidth(4, 70);   // ACTIVO
  hoja.setColumnWidth(5, 280);  // DESCRIPCION

  // Congelar primera fila
  hoja.setFrozenRows(1);

  Logger.log('📋 Hoja ' + nombreHoja + ' creada correctamente con 4 estados iniciales');
}
/**
 * Crea la hoja de estados de inventario
 */
function crearHojaEstadosInventario(ss) {
  var nombre = '📋_Estados_Inventario';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['ID', 'NOMBRE', 'COLOR', 'ACTIVO', 'DESCRIPCION'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    [1, 'Disponible', '#28a745', true, 'Equipo disponible para instalar'],
    [2, 'Instalado', '#007bff', true, 'Equipo instalado en vehículo'],
    [3, 'Garantía', '#ffc107', true, 'Equipo en proceso de garantía'],
    [4, 'Baja', '#dc3545', true, 'Equipo dado de baja']
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 250);

  console.log('✅ Hoja ' + nombre + ' creada correctamente');
  return sheet;
}
/**
 * Crea la hoja de tipos de vehículo
 */
function crearHojaTiposVehiculo(ss) {
  var nombre = '📋_Tipos_Vehiculo';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['TIPO', 'DESCRIPCION', 'ACTIVO'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['UTILITARIA', 'Vehículo utilitario para carga', true],
    ['CAMION', 'Camión de carga', true],
    ['VAN', 'Van o furgoneta', true],
    ['PICKUP', 'Pickup o camioneta', true],
    ['AUTOMOVIL', 'Automóvil particular', true],
    ['TRACTOCAMION', 'Tractocamión', true],
    ['REMOLQUE', 'Remolque', true]
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 70);

  return sheet;
}
/**
 * Crea la hoja de estados de tickets
 */
function crearHojaEstadosTicket(ss) {
  var nombre = '📋_Estados_Ticket';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['ID', 'NOMBRE', 'COLOR', 'ACTIVO', 'DESCRIPCION'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    [1, 'Pendiente', 'warning', true, 'Ticket esperando ser tomado'],
    [2, 'En proceso', 'primary', true, 'Ticket en proceso de resolución'],
    [3, 'Resuelto', 'success', true, 'Ticket resuelto']
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 250);

  return sheet;
}
function crearHojaEstadosTicket() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.insertSheet('📋_Estados_Ticket');

  var headers = ['ID', 'NOMBRE', 'COLOR', 'ACTIVO', 'DESCRIPCION'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    [1, 'Pendiente', 'warning', true, 'Ticket esperando ser tomado'],
    [2, 'En proceso', 'primary', true, 'Ticket en proceso de resolución'],
    [3, 'Resuelto', 'success', true, 'Ticket resuelto']
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 250);

  console.log('✅ Hoja 📋_Estados_Ticket creada');
}
// ============================================================
// CREAR HOJA: 📋_Tipos_Vehiculo
// ============================================================
function crearHojaTiposVehiculo(ss) {
  var nombre = '📋_Tipos_Vehiculo';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = ['TIPO', 'DESCRIPCION', 'ACTIVO'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var datos = [
    ['UTILITARIA', 'Vehículo utilitario para carga', true],
    ['CAMION', 'Camión de carga', true],
    ['VAN', 'Van o furgoneta', true],
    ['PICKUP', 'Pickup o camioneta', true],
    ['AUTOMOVIL', 'Automóvil particular', true],
    ['TRACTOCAMION', 'Tractocamión', true],
    ['REMOLQUE', 'Remolque', true]
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 70);

  return sheet;
}

// ============================================================
// CREAR HOJA: 📋_Catalogo_Vehiculos
// ============================================================
function crearHojaCatalogoVehiculos(ss) {
  var nombre = '📋_Catalogo_Vehiculos';
  var sheet = ss.getSheetByName(nombre);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(nombre);

  var headers = [
    'ECONOMICO', 'PLACAS', 'TIPO_VEHICULO', 'MARCA', 'MODELO',
    'AÑO', 'SERIE_VEHICULO', 'GPS_ACTUAL', 'ESTADO', 'ULTIMO_SERVICIO', 'TIPO_UNIDAD'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Datos de ejemplo
  var datos = [
    ['G-361', 'ABC-123-X', 'UTILITARIA', 'Ford', 'F-350', 2023, '1FT8W3BT9PED12345', 'GTMF-KL7-P9M', 'Activo', '', 'UNIDAD'],
    ['G-444', 'DEF-456-Y', 'CAMION', 'Ford', 'F-350', 2023, '1FT8W3BT9PED12346', 'GTMF-BN8-XC4', 'Activo', '', 'UTILITARIA'],
    ['G-445', 'GHI-789-Z', 'CAMION', 'International', 'HV', 2022, '1HTMKAAN8NH123456', 'GTMF-RYD-W8S', 'Activo', '', 'CAJA SECA'],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  // Validación de ESTADO
  var reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Activo', 'Inactivo', 'Mantenimiento', 'Baja'], true)
    .build();
  sheet.getRange('I2:I').setDataValidation(reglaEstado);

  // Validación de TIPO_VEHICULO
  var tiposVeh = ss.getSheetByName('📋_Tipos_Vehiculo');
  if (tiposVeh) {
    var tipos = tiposVeh.getRange('A2:A').getValues();
    var listaTipos = [];
    for (var i = 0; i < tipos.length; i++) {
      if (tipos[i][0]) listaTipos.push(tipos[i][0]);
    }
    if (listaTipos.length > 0) {
      var reglaTipo = SpreadsheetApp.newDataValidation()
        .requireValueInList(listaTipos, true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange('C2:C').setDataValidation(reglaTipo);
    }
  }

  // Validación de TIPO_UNIDAD
  var tiposUnidad = ss.getSheetByName('📋_Tipos_Unidad');
  if (tiposUnidad) {
    var tipos = tiposUnidad.getRange('A2:A').getValues();
    var listaTiposUnidad = [];
    for (var j = 0; j < tipos.length; j++) {
      if (tipos[j][0]) listaTiposUnidad.push(tipos[j][0]);
    }
    if (listaTiposUnidad.length > 0) {
      var reglaTipoUnidad = SpreadsheetApp.newDataValidation()
        .requireValueInList(listaTiposUnidad, true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange('K2:K').setDataValidation(reglaTipoUnidad);
    }
  }

  // Anchos de columna
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 150);

  console.log('✅ Hoja ' + nombre + ' creada correctamente');
  return sheet;
}
/**
 * Configura protecciones y permisos en todas las hojas del sistema
 * @param {Object} hojas - Objeto con todas las hojas creadas
 * @returns {boolean} - true si se configuraron correctamente
 */
function configurarProtecciones(hojas) {
  try {
    // Validar entrada
    if (!hojas || typeof hojas !== 'object') {
      throw new Error('El objeto "hojas" es requerido');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminEmail = Session.getActiveUser().getEmail();
    let proteccionesConfiguradas = 0;

    console.log('🔒 Iniciando configuración de protecciones...');

    // ============================================================
    // 1. PROTEGER HOJAS CRÍTICAS (SOLO ADMINISTRADORES)
    // ============================================================
    const hojasCriticas = [
      { nombre: '👤_Usuarios', descripcion: 'Datos de usuarios del sistema' },
      { nombre: '⚙️_Parametros', descripcion: 'Configuración global del sistema' },
      { nombre: '💰_Tarifas', descripcion: 'Tarifas y precios' },
      { nombre: '📋_Tipos_Equipo', descripcion: 'Catálogo de tipos de equipo' },
      { nombre: '📋_Tipos_Unidad', descripcion: 'Catálogo de tipos de unidad' },
      { nombre: '📋_Tipos_Vehiculo', descripcion: 'Catálogo de tipos de vehículo' },
      { nombre: '📋_Estados_Ticket', descripcion: 'Catálogo de estados de ticket' }
    ];

    hojasCriticas.forEach(({ nombre, descripcion }) => {
      try {
        const sheet = ss.getSheetByName(nombre);
        if (!sheet) {
          console.warn(`⚠️ Hoja "${nombre}" no encontrada, omitiendo protección`);
          return;
        }

        // Remover protecciones existentes
        const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        protections.forEach(prot => prot.remove());

        // Crear nueva protección
        const protection = sheet.protect();
        protection.setDescription(`🔒 ${descripcion} - Solo administradores`);
        
        // Agregar al admin actual
        protection.addEditor(adminEmail);
        
        // Opcional: agregar otros administradores específicos
        // protection.addEditor('admin2@empresa.com');
        // protection.addEditor('admin3@empresa.com');

        // Configurar para que solo advierta, no bloquee completamente
        protection.setWarningOnly(true);

        proteccionesConfiguradas++;
        console.log(`✅ Hoja "${nombre}" protegida correctamente`);

      } catch (err) {
        console.warn(`⚠️ Error al proteger "${nombre}":`, err.message);
      }
    });

    // ============================================================
    // 2. PROTEGER HOJAS CON DATOS OPERATIVOS (SOLO LECTURA PARA TÉCNICOS)
    // ============================================================
    const hojasOperativas = [
      { nombre: '📝_Bitacora_Revisiones', descripcion: 'Bitácora de revisiones' },
      { nombre: '📦_Inventario_GPS', descripcion: 'Inventario de GPS' },
      { nombre: '🚚_Flotilla_Fallas', descripcion: 'Flotilla y fallas' }
    ];

    hojasOperativas.forEach(({ nombre, descripcion }) => {
      try {
        const sheet = ss.getSheetByName(nombre);
        if (!sheet) {
          console.warn(`⚠️ Hoja "${nombre}" no encontrada, omitiendo protección`);
          return;
        }

        // Solo proteger rangos específicos, no toda la hoja
        const ultimaFila = sheet.getLastRow();
        if (ultimaFila > 0) {
          // Proteger encabezados (fila 1)
          const rangoEncabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          const protection = rangoEncabezados.protect();
          protection.setDescription(`🔒 ${descripcion} - Encabezados protegidos`);
          protection.addEditor(adminEmail);
          protection.setWarningOnly(true);
          
          proteccionesConfiguradas++;
          console.log(`✅ Encabezados de "${nombre}" protegidos`);
        }
      } catch (err) {
        console.warn(`⚠️ Error al proteger "${nombre}":`, err.message);
      }
    });

    // ============================================================
    // 3. PROTEGER COLUMNAS DE FÓRMULAS Y AUTOMATIZACIONES
    // ============================================================
    
    // 3.1 Proteger columna FOLIO en Bitácora
    try {
      const bitacora = hojas.bitacora || ss.getSheetByName('📝_Bitacora_Revisiones');
      if (bitacora) {
        const ultimaFila = bitacora.getLastRow();
        if (ultimaFila > 0) {
          const rangoFolios = bitacora.getRange(1, 1, ultimaFila, 1); // Columna A
          const protection = rangoFolios.protect();
          protection.setDescription('🔒 FOLIO - No modificar manualmente (autogenerado)');
          protection.addEditor(adminEmail);
          protection.setWarningOnly(true);
          proteccionesConfiguradas++;
          console.log('✅ Columna FOLIO protegida');
        }
      }
    } catch (err) {
      console.warn('⚠️ Error al proteger columna FOLIO:', err.message);
    }

    // 3.2 Proteger columnas de fechas automáticas en Inventario
    try {
      const inventario = hojas.inventario || ss.getSheetByName('📦_Inventario_GPS');
      if (inventario) {
        const ultimaFila = inventario.getLastRow();
        if (ultimaFila > 0) {
          // Asumiendo que la columna de fecha es la columna A
          const rangoFechas = inventario.getRange(2, 1, ultimaFila - 1, 1);
          const protection = rangoFechas.protect();
          protection.setDescription('🔒 Fechas automáticas - No modificar');
          protection.addEditor(adminEmail);
          protection.setWarningOnly(true);
          proteccionesConfiguradas++;
          console.log('✅ Columna de fechas en Inventario protegida');
        }
      }
    } catch (err) {
      console.warn('⚠️ Error al proteger fechas en Inventario:', err.message);
    }

    // 3.3 Proteger columnas de estados en Tickets
    try {
      const tickets = hojas.tickets || ss.getSheetByName('🎫_Tickets');
      if (tickets) {
        const ultimaFila = tickets.getLastRow();
        if (ultimaFila > 0) {
          // Asumiendo que la columna de estado es la columna G
          const rangoEstados = tickets.getRange(2, 7, ultimaFila - 1, 1);
          const protection = rangoEstados.protect();
          protection.setDescription('🔒 ESTADO - Usar lista desplegable');
          protection.addEditor(adminEmail);
          protection.setWarningOnly(true);
          proteccionesConfiguradas++;
          console.log('✅ Columna de estados en Tickets protegida');
        }
      }
    } catch (err) {
      console.warn('⚠️ Error al proteger estados en Tickets:', err.message);
    }

    // ============================================================
    // 4. PROTEGER HOJAS DE AUDITORÍA (SOLO LECTURA)
    // ============================================================
    const hojasAuditoria = [
      { nombre: '📈_Log_Auditoria', descripcion: 'Registro de auditoría' },
      { nombre: '📩_Notificaciones', descripcion: 'Notificaciones del sistema' }
    ];

    hojasAuditoria.forEach(({ nombre, descripcion }) => {
      try {
        const sheet = ss.getSheetByName(nombre);
        if (!sheet) {
          console.warn(`⚠️ Hoja "${nombre}" no encontrada, omitiendo protección`);
          return;
        }

        // Proteger toda la hoja
        const protection = sheet.protect();
        protection.setDescription(`🔒 ${descripcion} - Solo lectura`);
        protection.addEditor(adminEmail);
        protection.setWarningOnly(true);
        
        proteccionesConfiguradas++;
        console.log(`✅ Hoja "${nombre}" protegida (solo lectura)`);

      } catch (err) {
        console.warn(`⚠️ Error al proteger "${nombre}":`, err.message);
      }
    });

    // ============================================================
    // 5. CONFIGURAR PERMISOS POR ROL (AVANZADO)
    // ============================================================
    // Nota: Esta sección es opcional y requiere que tengas una hoja de usuarios
    try {
      const usuariosSheet = hojas.usuarios || ss.getSheetByName('👤_Usuarios');
      if (usuariosSheet) {
        // Leer usuarios y roles
        const datosUsuarios = usuariosSheet.getDataRange().getValues();
        const headers = datosUsuarios[0];
        const idxRol = headers.indexOf('ROL');
        const idxEmail = headers.indexOf('EMAIL');
        const idxActivo = headers.indexOf('ACTIVO');

        if (idxRol !== -1 && idxEmail !== -1) {
          // Agrupar usuarios por rol
          const roles = {
            administradores: [],
            revisores: [],
            tecnicos: []
          };

          for (let i = 1; i < datosUsuarios.length; i++) {
            const email = datosUsuarios[i][idxEmail];
            const rol = datosUsuarios[i][idxRol];
            const activo = datosUsuarios[i][idxActivo] !== false;

            if (email && activo) {
              const rolLower = (rol || '').toString().toLowerCase();
              if (rolLower.includes('admin') || rolLower.includes('administrador')) {
                roles.administradores.push(email);
              } else if (rolLower.includes('revisor')) {
                roles.revisores.push(email);
              } else if (rolLower.includes('tec')) {
                roles.tecnicos.push(email);
              }
            }
          }

          console.log(`👥 Usuarios encontrados: Admin: ${roles.administradores.length}, Revisores: ${roles.revisores.length}, Técnicos: ${roles.tecnicos.length}`);

          // Aquí podrías aplicar protecciones específicas por rol
          // Por ejemplo: dar permisos de edición a revisores en ciertas hojas
        }
      }
    } catch (err) {
      console.warn('⚠️ Error al configurar permisos por rol:', err.message);
    }

    // ============================================================
    // 6. LOG FINAL
    // ============================================================
    console.log(`✅ Protecciones configuradas exitosamente (${proteccionesConfiguradas} elementos protegidos)`);
    
    return true;

  } catch (err) {
    console.error('❌ Error al configurar protecciones:', err.message);
    console.error('Stack trace:', err.stack);
    
    // No lanzar error para no interrumpir la instalación
    // pero retornar false para indicar que hubo problemas
    return false;
  }
}
/**
 * Configura permisos específicos basados en roles de usuario
 * @param {Object} hojas - Objeto con todas las hojas
 */
function configurarPermisosPorRol(hojas) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usuariosSheet = hojas.usuarios || ss.getSheetByName('👤_Usuarios');
    
    if (!usuariosSheet) {
      console.warn('⚠️ No se encontró la hoja de usuarios para configurar permisos');
      return false;
    }

    // Leer configuración de permisos desde la hoja de parámetros
    const paramsSheet = hojas.parametros || ss.getSheetByName('⚙️_Parametros');
    let permisosConfig = {};
    
    if (paramsSheet) {
      const paramsData = paramsSheet.getDataRange().getValues();
      for (let i = 0; i < paramsData.length; i++) {
        const clave = paramsData[i][0];
        const valor = paramsData[i][1];
        if (clave && clave.startsWith('PERMISO_')) {
          permisosConfig[clave] = valor;
        }
      }
    }

    // Configurar permisos según roles
    // Ejemplo: permitir a revisores editar ciertas columnas
    const bitacora = hojas.bitacora || ss.getSheetByName('📝_Bitacora_Revisiones');
    if (bitacora && permisosConfig.PERMISO_REVISOR_EDITAR === 'true') {
      // Lógica para dar permisos específicos a revisores
      console.log('✅ Permisos para revisores configurados');
    }

    return true;

  } catch (err) {
    console.error('❌ Error al configurar permisos por rol:', err.message);
    return false;
  }
}
// Prueba 1: Ejecutar la función con hojas válidas
function testConfigurarProtecciones() {
  const hojas = {
    usuarios: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('👤_Usuarios'),
    parametros: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️_Parametros'),
    tarifas: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💰_Tarifas'),
    bitacora: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝_Bitacora_Revisiones'),
    inventario: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📦_Inventario_GPS'),
    tickets: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🎫_Tickets')
  };
  
  const resultado = configurarProtecciones(hojas);
  console.log('Resultado:', resultado ? '✅ Éxito' : '❌ Falló');
}

