/**
 * ============================================================
 * SISTEMA DE GESTIÓN DE NOVEDADES - POLICÍA NACIONAL DE COLOMBIA
 * Google Apps Script Backend
 * ============================================================
 * 
 * Estructura de cada hoja (dependencia):
 * Columna A: ID (fila)
 * Columna B: GRADO
 * Columna C: (vacía o datos adicionales)
 * Columna D: (vacía o datos adicionales)
 * Columna E: APELLIDOS Y NOMBRES
 * Columna F: TURNO
 * Columna G: NOVEDAD (lista desplegable)
 * Columna H: OBSERVACION
 * 
 * Fila 1: Encabezados
 * Filas 2+: Datos del personal
 */

// ============================================================
// CONFIGURACIÓN
// ============================================================
const CONFIG = {
  // Hoja de usuarios (nombre fijo)
  USERS_SHEET: 'USUARIOS',
  
  // Fila de encabezados
  HEADER_ROW: 1,
  
  // Columnas (índices 1-based para getRange)
  COL_ID: 1,        // A
  COL_GRADO: 2,     // B
  COL_NOMBRES: 5,   // E
  COL_TURNO: 6,     // F
  COL_NOVEDAD: 7,   // G
  COL_OBS: 8,       // H
  
  // CORS Headers
  CORS_HEADERS: {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  }
};

// ============================================================
// MANEJO DE CORS - doOptions (obligatorio para POST desde frontend)
// ============================================================
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders(CONFIG.CORS_HEADERS);
}

// ============================================================
// doGet - Maneja todas las solicitudes GET
// ============================================================
function doGet(e) {
  const action = e.parameter.action || 'getData';
  const dependency = e.parameter.dependency || 'ASJUR';
  
  try {
    let result;
    
    switch(action) {
      case 'login':
        result = handleLogin(e.parameter.username, e.parameter.password);
        break;
      case 'verifyDepKey':
        result = handleVerifyDepKey(e.parameter.username, e.parameter.depKey, e.parameter.dependencia);
        break;
      case 'getData':
        result = handleGetData(dependency);
        break;
      case 'getSummary':
        result = handleGetSummary(dependency, e.parameter.turno);
        break;
      case 'getTotalFuncionarios':
        result = handleGetTotalFuncionarios(dependency);
        break;
      default:
        result = { success: false, error: 'Acción no válida: ' + action };
    }
    
    return createJSONResponse(result);
    
  } catch (error) {
    console.error('Error en doGet:', error);
    return createJSONResponse({ success: false, error: error.toString() });
  }
}

// ============================================================
// doPost - Maneja todas las solicitudes POST (guardar datos)
// ============================================================
function doPost(e) {
  try {
    const action = e.parameter.action || '';
    const dependency = e.parameter.dependency || 'ASJUR';
    
    console.log('doPost recibido:', { action: action, dependency: dependency, params: e.parameter });
    
    let result;
    
    switch(action) {
      case 'saveData':
        result = handleSaveData(dependency, e.parameter);
        break;
      case 'login':
        result = handleLogin(e.parameter.username, e.parameter.password);
        break;
      case 'verifyDepKey':
        result = handleVerifyDepKey(e.parameter.username, e.parameter.depKey, e.parameter.dependencia);
        break;
      default:
        result = { success: false, error: 'Acción POST no válida: ' + action };
    }
    
    return createJSONResponse(result);
    
  } catch (error) {
    console.error('Error en doPost:', error);
    return createJSONResponse({ success: false, error: error.toString() });
  }
}

// ============================================================
// HELPERS - Crear respuesta JSON con CORS
// ============================================================
function createJSONResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  
  // Agregar headers CORS
  for (const key in CONFIG.CORS_HEADERS) {
    output.setHeader(key, CONFIG.CORS_HEADERS[key]);
  }
  
  return output;
}

function createJSONPResponse(data, callback) {
  const jsonString = JSON.stringify(data);
  const output = ContentService.createTextOutput(callback + '(' + jsonString + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  
  // Agregar headers CORS
  for (const key in CONFIG.CORS_HEADERS) {
    output.setHeader(key, CONFIG.CORS_HEADERS[key]);
  }
  
  return output;
}

// ============================================================
// LOGIN - Verificar usuario y contraseña
// ============================================================
function handleLogin(username, password) {
  if (!username || !password) {
    return { success: false, error: 'Usuario y contraseña requeridos' };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
    
    if (!sheet) {
      return { success: false, error: 'Sistema de usuarios no configurado' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Buscar usuario (fila 1 = headers, desde fila 2)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUsername = (row[0] || '').toString().trim();
      const rowPassword = (row[1] || '').toString().trim();
      const rowRol = (row[2] || '').toString().trim();
      const rowDependencia = (row[3] || '').toString().trim();
      const rowToken = (row[4] || '').toString().trim();
      
      if (rowUsername === username && rowPassword === password) {
        return {
          success: true,
          usuario: {
            username: rowUsername,
            rol: rowRol,
            dependencia: rowDependencia,
            token: rowToken || Date.now().toString()
          }
        };
      }
    }
    
    return { success: false, error: 'Usuario o contraseña incorrectos' };
    
  } catch (error) {
    console.error('Error login:', error);
    return { success: false, error: 'Error del servidor: ' + error.toString() };
  }
}

// ============================================================
// VERIFICAR CLAVE DE DEPENDENCIA
// ============================================================
function handleVerifyDepKey(username, depKey, dependencia) {
  if (!username || !depKey || !dependencia) {
    return { success: false, error: 'Datos incompletos' };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Buscar en la hoja de dependencias o en la hoja del usuario
    // Aquí verificamos que la clave coincida con la dependencia
    const sheet = ss.getSheetByName(dependencia);
    
    if (!sheet) {
      return { success: false, error: 'Dependencia no encontrada' };
    }
    
    // Verificar clave en la primera celda o en una celda específica
    // Por defecto, asumimos que la clave está en la celda J1 de cada hoja
    const claveGuardada = sheet.getRange('J1').getValue().toString().trim();
    
    if (claveGuardada === depKey) {
      return { success: true };
    } else {
      return { success: false, error: 'Clave de dependencia incorrecta' };
    }
    
  } catch (error) {
    console.error('Error verifyDepKey:', error);
    return { success: false, error: 'Error del servidor: ' + error.toString() };
  }
}

// ============================================================
// GET DATA - Obtener todos los funcionarios de una dependencia
// ============================================================
function handleGetData(dependency) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    
    if (!sheet) {
      return { success: false, error: 'Dependencia no encontrada: ' + dependency };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    // Leer datos desde fila 2 (después del header)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 8);
    const values = dataRange.getValues();
    
    const result = [];
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowNum = i + 2; // Fila real en Sheets (2, 3, 4...)
      
      result.push({
        row: rowNum,                    // Número de fila real para guardar después
        grado: row[CONFIG.COL_GRADO - 1] || '',
        nombres: row[CONFIG.COL_NOMBRES - 1] || '',
        turno: (row[CONFIG.COL_TURNO - 1] || '').toString().trim().toUpperCase(),
        novedad: row[CONFIG.COL_NOVEDAD - 1] || '',
        observacion: row[CONFIG.COL_OBS - 1] || ''
      });
    }
    
    return { success: true, data: result };
    
  } catch (error) {
    console.error('Error getData:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// SAVE DATA - Guardar novedad y observación
// ============================================================
function handleSaveData(dependency, params) {
  const row = parseInt(params.row);
  const novedad = params.novedad || '';
  const observation = params.observation || '';
  
  console.log('handleSaveData:', { dependency: dependency, row: row, novedad: novedad, observation: observation });
  
  if (!row || isNaN(row) || row < 2) {
    return { success: false, error: 'Fila inválida: ' + params.row };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    
    if (!sheet) {
      return { success: false, error: 'Dependencia no encontrada: ' + dependency };
    }
    
    // Verificar que la fila existe
    const lastRow = sheet.getLastRow();
    if (row > lastRow) {
      return { success: false, error: 'Fila ' + row + ' no existe. Última fila: ' + lastRow };
    }
    
    // Guardar novedad en columna G
    sheet.getRange(row, CONFIG.COL_NOVEDAD).setValue(novedad);
    
    // Guardar observación en columna H
    sheet.getRange(row, CONFIG.COL_OBS).setValue(observation);
    
    // Agregar timestamp en columna I (opcional, para auditoría)
    // sheet.getRange(row, 9).setValue(new Date());
    
    console.log('Datos guardados exitosamente en fila', row);
    
    return { 
      success: true, 
      message: 'Datos guardados correctamente',
      row: row,
      novedad: novedad,
      observation: observation
    };
    
  } catch (error) {
    console.error('Error saveData:', error);
    return { success: false, error: 'Error al guardar: ' + error.toString() };
  }
}

// ============================================================
// GET TOTAL FUNCIONARIOS
// ============================================================
function handleGetTotalFuncionarios(dependency) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    
    if (!sheet) {
      return { success: false, error: 'Dependencia no encontrada' };
    }
    
    const lastRow = sheet.getLastRow();
    const total = lastRow > 1 ? lastRow - 1 : 0;
    
    return { success: true, total: total };
    
  } catch (error) {
    console.error('Error getTotalFuncionarios:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GET SUMMARY - Resumen por turno
// ============================================================
function handleGetSummary(dependency, turno) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    
    if (!sheet) {
      return { success: false, error: 'Dependencia no encontrada' };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, summary: [] };
    }
    
    // Leer todas las filas
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 8);
    const values = dataRange.getValues();
    
    // Filtrar por turno si se especifica
    let filteredData = values;
    if (turno) {
      filteredData = values.filter(row => {
        const t = (row[CONFIG.COL_TURNO - 1] || '').toString().trim().toUpperCase();
        return t === turno.toUpperCase();
      });
    }
    
    // Contar totales
    const total = filteredData.length;
    const conNovedad = filteredData.filter(r => (r[CONFIG.COL_NOVEDAD - 1] || '').toString().trim() !== '').length;
    const sinNovedad = total - conNovedad;
    
    // Contar por tipo de novedad
    const novedadesCount = {};
    filteredData.forEach(row => {
      const nov = (row[CONFIG.COL_NOVEDAD - 1] || '').toString().trim();
      if (nov) {
        novedadesCount[nov] = (novedadesCount[nov] || 0) + 1;
      }
    });
    
    // Construir respuesta de resumen
    const summary = [];
    
    // Header del resumen
    summary.push({
      section: 'header',
      label: 'RESUMEN ' + dependency + (turno ? ' - TURNO ' + turno : ''),
      value: ''
    });
    
    // Totales
    summary.push({ section: 'normal', label: 'Total Funcionarios', value: total });
    summary.push({ section: 'normal', label: 'Con Novedad', value: conNovedad });
    summary.push({ section: 'normal', label: 'Sin Novedad', value: sinNovedad });
    
    // Novedades detalladas
    if (Object.keys(novedadesCount).length > 0) {
      summary.push({ section: 'novedad-header', label: 'DETALLE DE NOVEDADES', value: '' });
      
      for (const [novedad, count] of Object.entries(novedadesCount)) {
        summary.push({
          section: 'novedad-item',
          label: novedad,
          value: count
        });
      }
    }
    
    return { success: true, summary: summary };
    
  } catch (error) {
    console.error('Error getSummary:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// FUNCIONES AUXILIARES PARA ADMINISTRADOR
// ============================================================

/**
 * Crear una nueva hoja de dependencia con estructura inicial
 * @param {string} nombreDependencia - Nombre de la dependencia (nombre de la hoja)
 * @param {string} claveDependencia - Clave de acceso para la dependencia
 */
function crearHojaDependencia(nombreDependencia, claveDependencia) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Verificar si ya existe
  let sheet = ss.getSheetByName(nombreDependencia);
  if (sheet) {
    return 'La hoja ' + nombreDependencia + ' ya existe';
  }
  
  // Crear nueva hoja
  sheet = ss.insertSheet(nombreDependencia);
  
  // Encabezados
  const headers = ['ID', 'GRADO', '', '', 'APELLIDOS Y NOMBRES', 'TURNO', 'NOVEDAD', 'OBSERVACION'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formato de encabezados
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#01592F')
    .setFontColor('#FFFFFF');
  
  // Ajustar anchos de columna
  sheet.setColumnWidth(1, 50);   // ID
  sheet.setColumnWidth(2, 80);   // GRADO
  sheet.setColumnWidth(5, 250);  // NOMBRES
  sheet.setColumnWidth(6, 60);   // TURNO
  sheet.setColumnWidth(7, 150);  // NOVEDAD
  sheet.setColumnWidth(8, 200);  // OBSERVACION
  
  // Validación de datos para columna G (NOVEDAD) - lista desplegable
  const novedades = [
    'S/N', 'OFICINA', 'VACACIONES', 'SERVICIO', 'ESTADO DE GRAVIDEZ',
    'EXCUSA TOTAL', 'PERMISO', 'CITA MEDICA', 'LICENCIA DE MATERNIDAD',
    'EXCUSADO DEL SERVICIO', 'OTRA NOVEDAD', 'RETARDAD@', 'EXCUSA PARCIAL',
    'CURSO DE ASCENSO', 'COMISION'
  ];
  
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(novedades, true)
    .setAllowInvalid(false)
    .setHelpText('Seleccione una novedad válida')
    .build();
  
  // Aplicar validación a toda la columna G (desde G2 hasta G1000)
  sheet.getRange('G2:G1000').setDataValidation(rule);
  
  // Guardar clave en celda J1 (oculta, para verificación)
  sheet.getRange('J1').setValue(claveDependencia);
  sheet.getRange('J1').setFontColor('#FFFFFF'); // Texto blanco (oculto visualmente)
  
  // Congelar fila de encabezados
  sheet.setFrozenRows(1);
  
  return 'Hoja ' + nombreDependencia + ' creada exitosamente';
}

/**
 * Crear hoja de usuarios con estructura inicial
 */
function crearHojaUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (sheet) {
    return 'La hoja USUARIOS ya existe';
  }
  
  sheet = ss.insertSheet(CONFIG.USERS_SHEET);
  
  // Encabezados: USUARIO | CONTRASEÑA | ROL | DEPENDENCIA | TOKEN
  const headers = ['USUARIO', 'CONTRASEÑA', 'ROL', 'DEPENDENCIA', 'TOKEN'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#01592F')
    .setFontColor('#FFFFFF');
  
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 150);
  
  return 'Hoja USUARIOS creada exitosamente';
}

/**
 * Agregar usuario
 */
function agregarUsuario(username, password, rol, dependencia) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  
  if (!sheet) {
    return 'Error: Hoja USUARIOS no existe. Ejecute crearHojaUsuarios() primero.';
  }
  
  const token = Utilities.getUuid();
  const lastRow = sheet.getLastRow();
  
  sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[username, password, rol, dependencia, token]]);
  
  return 'Usuario ' + username + ' agregado exitosamente';
}

/**
 * Función para ejecutar desde el editor (menú)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚔 Policía Nacional')
    .addItem('➕ Crear Hoja de Dependencia', 'mostrarDialogoCrearDependencia')
    .addItem('👤 Crear Hoja de Usuarios', 'crearHojaUsuarios')
    .addItem('➕ Agregar Usuario', 'mostrarDialogoAgregarUsuario')
    .addSeparator()
    .addItem('📋 Ver Dependencias', 'listarDependencias')
    .addToUi();
}

function mostrarDialogoCrearDependencia() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      input { width: 100%; padding: 8px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; }
      button { background: #01592F; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
      button:hover { background: #013d1f; }
    </style>
    <h3>Crear Nueva Dependencia</h3>
    <input type="text" id="nombre" placeholder="Nombre de la dependencia (ej: ASJUR)">
    <input type="text" id="clave" placeholder="Clave de acceso">
    <button onclick="crear()">Crear</button>
    <div id="resultado"></div>
    <script>
      function crear() {
        const nombre = document.getElementById('nombre').value.trim();
        const clave = document.getElementById('clave').value.trim();
        if (!nombre || !clave) { alert('Complete todos los campos'); return; }
        google.script.run.withSuccessHandler(function(r) {
          document.getElementById('resultado').innerHTML = '<p style="color:green">' + r + '</p>';
        }).crearHojaDependencia(nombre, clave);
      }
    </script>
  `)
  .setWidth(350)
  .setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Crear Dependencia');
}

function mostrarDialogoAgregarUsuario() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      input, select { width: 100%; padding: 8px; margin: 5px 0; border: 1px solid #ccc; border-radius: 4px; }
      button { background: #01592F; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
      button:hover { background: #013d1f; }
    </style>
    <h3>Agregar Usuario</h3>
    <input type="text" id="username" placeholder="Usuario">
    <input type="password" id="password" placeholder="Contraseña">
    <select id="rol">
      <option value="usuario">Usuario</option>
      <option value="administrador">Administrador</option>
    </select>
    <input type="text" id="dependencia" placeholder="Dependencia (ej: ASJUR)">
    <button onclick="agregar()">Agregar</button>
    <div id="resultado"></div>
    <script>
      function agregar() {
        const u = document.getElementById('username').value.trim();
        const p = document.getElementById('password').value.trim();
        const r = document.getElementById('rol').value;
        const d = document.getElementById('dependencia').value.trim();
        if (!u || !p || !d) { alert('Complete todos los campos'); return; }
        google.script.run.withSuccessHandler(function(res) {
          document.getElementById('resultado').innerHTML = '<p style="color:green">' + res + '</p>';
        }).agregarUsuario(u, p, r, d);
      }
    </script>
  `)
  .setWidth(350)
  .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Usuario');
}

function listarDependencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const nombres = sheets.map(s => s.getName()).filter(n => n !== CONFIG.USERS_SHEET);
  
  SpreadsheetApp.getUi().alert('Dependencias existentes: ' + nombres.join(', '));
}

// ============================================================
// FUNCIÓN DE PRUEBA
// ============================================================
function test() {
  console.log('Test ejecutado correctamente');
  console.log('Hojas disponibles:', SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()));
}