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
function instalarBaseDatos() {
  try {
    // 🔍 DETECTAR CONTEXTO DE EJECUCIÓN
    var esEditor = false;
    var ui = null;
    
    try {
      // Intentar obtener UI - SOLO disponible en editor o menú
      ui = SpreadsheetApp.getUi();
      esEditor = true;
    } catch (e) {
      // No hay UI disponible (ejecución desde Web App o trigger)
      esEditor = false;
    }
    
    // ✅ Si NO es editor, mostrar error claro y detener
    if (!esEditor) {
      var mensaje = '❌ Esta función debe ejecutarse desde el EDITOR DE SCRIPTS.\n\n' +
                    'Pasos:\n' +
                    '1. Abre el editor: Extensiones → Apps Script\n' +
                    '2. Selecciona la función "instalarBaseDatos"\n' +
                    '3. Haz clic en "Ejecutar" (▶️)\n' +
                    '4. Autoriza los permisos si es necesario';
      
      // Intentar loguear el error (si hay consola disponible)
      console.error(mensaje);
      throw new Error(mensaje);
    }
    
    // ✅ Estamos en el editor, proceder con la instalación
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Confirmación (solo en editor)
    var respuesta = ui.alert(
      '⚠️ INSTALACIÓN DE BASE DE DATOS',
      'Este proceso eliminará TODAS las hojas existentes y creará una nueva estructura.\n' +
      '¿Estás seguro de que deseas continuar?',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta !== ui.Button.YES) {
      return;
    }
    
    // 2. Crear todas las hojas
    var hojas = {
      tarifas: crearHojatarifas(ss),
      usuarios: crearHojaUsuarios(ss),
      parametros: crearHojaParametros(ss),
      inventario: crearHojaInventario(ss),
      flotilla: crearHojaFlotilla(ss),
      bitacora: crearHojaBitacora(ss),
      consulta: crearHojaConsulta(ss),
      tiposEquipo: crearHojaTiposEquipo(ss),
      tiposUnidad: crearHojaTiposUnidad(ss),
      accesorios: crearHojaAccesorios(ss),
      facturas: crearHojaFacturas(ss),
      logAuditoria: crearHojaLogAuditoria(ss),
      notificaciones: crearHojaNotificaciones(ss),
      tickets: crearHojaTickets(ss),
      estadosVehiculo: crearHojaEstadosVehiculo(ss)
    };
    
    // 3. Configurar validaciones cruzadas
    configurarValidaciones(hojas);
    
    // 4. Configurar protecciones
    configurarProtecciones(hojas);
    
    // 5. Menú personalizado
    crearMenuPersonalizado();
    
    // 6. Mensaje de éxito
    mostrarMensajeExito(ui);
    
  } catch (err) {
    // Manejo de errores
    try {
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        '❌ ERROR',
        'Error durante la instalación:\n\n' + err.message,
        ui.ButtonSet.OK
      );
    } catch (e) {
      console.error('❌ Error en instalarBaseDatos:', err);
    }
    console.error('Error en instalarBaseDatos:', err);
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
  
  var headers = [
    'FECHA',
    'USUARIO',
    'FOLIO_AFECTADO',
    'CAMPO_MODIFICADO',
    'VALOR_ANTERIOR',
    'VALOR_NUEVO',
    'MOTIVO'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ANCHOS DE COLUMNA
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 250);
  
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
  
  var headers = [
    'ID', 'FECHA', 'UNIDAD', 'DESCRIPCION',
    'CREADO_POR', 'CREADO_POR_NOMBRE', 'ESTADO',
    'TECNICO_ASIGNADO', 'TECNICO_NOMBRE',
    'FECHA_CIERRE', 'COMENTARIOS', 'ULTIMA_ACTUALIZACION'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Datos de ejemplo
  var ahora = new Date();
  var hace2dias = new Date(ahora);
  hace2dias.setDate(hace2dias.getDate() - 2);
  var hace5dias = new Date(ahora);
  hace5dias.setDate(hace5dias.getDate() - 5);
  
  var datos = [
    ['TKT-001', hace5dias, 'G-361', 'El GPS no reporta posición. Posible falla en antena.', 'REV-001', 'María Valenzuela', 'Resuelto', 'TEC-001', 'Juan Pérez', ahora, 'Se reemplazó la antena GPS y se realizó prueba de funcionamiento. Todo OK.', ahora],
    ['TKT-002', hace2dias, 'G-444', 'La cámara no enciende. Se revisó cableado.', 'TEC-001', 'Juan Pérez', 'En proceso', 'TEC-001', 'Juan Pérez', null, 'Se detectó cable suelto en la conexión. Pendiente de soldadura.', ahora],
    ['TKT-003', new Date(), 'G-445', 'Falla en la conexión del gateway. No enciende.', 'TEC-002', 'Carlos Gómez', 'Pendiente', '', '', null, 'Se reporta el ticket para que un técnico lo revise.', ahora],
  ];
  sheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);
  
  // Validación de ESTADO
  var reglaEstadoTicket = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'En proceso', 'Resuelto'], true)
    .build();
  sheet.getRange('G2:G').setDataValidation(reglaEstadoTicket);
  
  return sheet;
}

// ============================================================
// 2. VALIDACIONES CRUZADAS
// ============================================================

function configurarValidaciones(hojas) {
  try {
    // 2.1 Validación de TIPO_REVISION en Bitácora contra Tarifas
    var bitacora = hojas.bitacora;
    var tarifas = hojas.tarifas;
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
    }
    
    // 2.2 Validación de ESTADO en Bitácora
    var reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Borrador', 'Listo para pago', 'Pagado'], true)
      .build();
    bitacora.getRange('M2:M').setDataValidation(reglaEstado);
    
    // 2.3 Validación de TIPO_EQUIPO en Inventario contra Tipos_Equipo
    var inventario = hojas.inventario;
    var tiposEquipo = hojas.tiposEquipo;
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
    
    // 2.4 Validación de TIPO_UNIDAD en Inventario contra Tipos_Unidad
    var tiposUnidad = hojas.tiposUnidad;
    if (tiposUnidad) {
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
  } catch (err) {
    console.warn('Error configurando validaciones:', err.message);
  }
}

// ============================================================
// 3. PROTECCIONES
// ============================================================

function configurarProtecciones(hojas) {
  try {
    var hojasProtegidas = ['💰_Tarifas', '👤_Usuarios', '⚙️_Parametros'];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    for (var i = 0; i < hojasProtegidas.length; i++) {
      var nombre = hojasProtegidas[i];
      var sheet = ss.getSheetByName(nombre);
      if (sheet) {
        var proteccion = sheet.protect();
        proteccion.setWarningOnly(true);
      }
    }
  } catch (err) {
    console.warn('No se pudieron configurar protecciones:', err.message);
  }
}

// ============================================================
// 4. MENÚ PERSONALIZADO
// ============================================================

function crearMenuPersonalizado() {
  try {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('📊 Fleet Manager');
    menu.addItem('🚀 Reinstalar base de datos', 'instalarBaseDatos');
    menu.addItem('📋 Cargar inventario desde Excel', 'cargarInventarioDesdeExcel');
    menu.addItem('🧹 Limpiar archivos Drive antiguos', 'limpiarDriveAntiguo');
    menu.addSeparator();
    menu.addItem('📊 Dashboard de técnicos', 'mostrarDashboard');
    menu.addItem('🔍 Verificar integridad', 'verificarIntegridad');
    menu.addToUi();
  } catch (err) {
    console.warn('No se pudo crear el menú:', err.message);
  }
}

// ============================================================
// 5. MENSAJE DE ÉXITO
// ============================================================

function mostrarMensajeExito(ui) {
  var mensaje = 
    '✅ INSTALACIÓN COMPLETA\n\n' +
    'Hojas creadas:\n' +
    '• 💰_Tarifas\n' +
    '• 👤_Usuarios\n' +
    '• ⚙️_Parametros\n' +
    '• 📦_Inventario_GPS (con TIPO_UNIDAD)\n' +
    '• 🚚_Flotilla_Fallas (con TIPO_UNIDAD)\n' +
    '• 📝_Bitacora_Revisiones\n' +
    '• 📊_Consulta_Tecnicos\n' +
    '• 📋_Tipos_Equipo\n' +
    '• 📋_Tipos_Unidad (NUEVA)\n' +
    '• 🔧_Accesorios_Stock\n' +
    '• 📑_Facturas\n' +
    '• 📈_Log_Auditoria\n' +
    '• 📩_Notificaciones\n' +
    '• 🎫_Tickets (NUEVA - Sistema de tickets/fallas)\n\n' +
    'USUARIOS DE PRUEBA:\n' +
    '• admin / admin123 (Administrador)\n' +
    '• revisor / rev123 (Revisor)\n' +
    '• tec1 / tec123 (Técnico)\n' +
    '• tec2 / tec123 (Técnico)';
  
  ui.alert('✅ INSTALACIÓN COMPLETA', mensaje, ui.ButtonSet.OK);
}

// ============================================================
// 6. UTILIDADES
// ============================================================

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

// ============================================================
// 7. FUNCIONES PLACEHOLDER
// ============================================================

function cargarInventarioDesdeExcel() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '📋 Carga de inventario',
    'Esta funcionalidad estará disponible próximamente.',
    ui.ButtonSet.OK
  );
}

function limpiarDriveAntiguo() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '🧹 Limpieza de Drive',
    'Esta funcionalidad estará disponible próximamente.',
    ui.ButtonSet.OK
  );
}

function mostrarDashboard() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '📊 Dashboard',
    'Esta funcionalidad estará disponible próximamente.',
    ui.ButtonSet.OK
  );
}

// ============================================================
// 8. VERIFICACIÓN DE INTEGRIDAD
// ============================================================

function verificarIntegridad() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojasEsperadas = [
    '💰_Tarifas',
    '👤_Usuarios',
    '⚙️_Parametros',
    '📦_Inventario_GPS',
    '🚚_Flotilla_Fallas',
    '📝_Bitacora_Revisiones',
    '📊_Consulta_Tecnicos',
    '📋_Tipos_Equipo',
    '📋_Tipos_Unidad',
    '🔧_Accesorios_Stock',
    '📑_Facturas',
    '📈_Log_Auditoria',
    '📩_Notificaciones',
    '🎫_Tickets'
  ];
  
  var resultados = [];
  for (var i = 0; i < hojasEsperadas.length; i++) {
    var nombre = hojasEsperadas[i];
    var exists = ss.getSheetByName(nombre) !== null;
    resultados.push((exists ? '✅' : '❌') + ' ' + nombre);
  }
  
  ui.alert(
    '🔍 VERIFICACIÓN DE INTEGRIDAD',
    'Resultados:\n\n' + resultados.join('\n'),
    ui.ButtonSet.OK
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