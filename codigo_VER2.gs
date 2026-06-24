''/**
 * ============================================================
 * SISTEMA DE GESTIÓN DE NOVEDADES - POLICÍA NACIONAL DE COLOMBIA
 * Google Apps Script Backend - VERSIÓN V7
 * 
 * CORRECCIÓN: GRADO está en columna D (índice 4), no en B (índice 2)
 * ============================================================
 */

const CONFIG = {
  USERS_SHEET: 'USUARIOS',
  HEADER_ROW: 1,
  COL_ID: 1,        // A
  COL_GRADO: 4,     // D  ← CORREGIDO: antes era 2 (B), ahora 4 (D)
  COL_NOMBRES: 5,   // E
  COL_TURNO: 6,     // F
  COL_NOVEDAD: 7,   // G
  COL_OBS: 8        // H
};

const SPREADSHEET_ID = '1NO1e1uov9EfrhaCJSmk7048ELPLxUkLGVjCDmgrlAbs';

// ============================================================
// doGet - Maneja TODAS las solicitudes (GET con JSONP)
// ============================================================
function doGet(e) {
  const action = e.parameter.action || '';
  const callback = e.parameter.callback || '';
  
  console.log('=== doGet === action=' + action + ' callback=' + (callback ? 'YES' : 'NO'));
  console.log('All params:', JSON.stringify(e.parameter));

  try {
    let result;

    switch (action) {
      case 'login':
        result = handleLogin(e.parameter.username, e.parameter.password);
        break;
      case 'verifyDepKey':
        result = handleVerifyDepKey(e.parameter.username, e.parameter.depKey, e.parameter.dependencia);
        break;
      case 'getData':
        result = handleGetData(e.parameter.dependency || '');
        break;
      case 'saveData':
        result = handleSaveData(e.parameter.dependency || '', e.parameter);
        break;
      case 'getSummary':
        result = handleGetSummary(e.parameter.dependency || '', e.parameter.turno);
        break;
      case 'getTotalFuncionarios':
        result = handleGetTotalFuncionarios(e.parameter.dependency || '');
        break;
      case 'getCells':
        result = handleGetCells(e.parameter.sheet, e.parameter.cells);
        break;
      case 'getAuditoria':
        result = handleGetAuditoria(e.parameter.dependency || '');
        break;
      case 'getUsuarios':
        result = handleGetUsuarios();
        break;
      default:
        result = { success: false, error: 'Acción no válida: ' + action };
    }

    console.log('Result:', JSON.stringify(result).substring(0, 300));

    if (callback && callback.trim() !== '') {
      return createJSONPResponse(result, callback);
    } else {
      return createJSONResponse(result);
    }

  } catch (error) {
    console.error('ERROR doGet:', error);
    const err = { success: false, error: error.toString() };
    if (callback && callback.trim() !== '') {
      return createJSONPResponse(err, callback);
    } else {
      return createJSONResponse(err);
    }
  }
}

// ============================================================
// HELPERS
// ============================================================
function createJSONResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function createJSONPResponse(data, callback) {
  const cb = callback.replace(/[^a-zA-Z0-9_]/g, '');
  const body = cb + '(' + JSON.stringify(data) + ');';
  console.log('JSONP body length:', body.length);
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ============================================================
// LOGIN
// ============================================================
function handleLogin(username, password) {
  console.log('=== handleLogin === user=' + (username || 'null'));

  if (!username || !password) {
    return { success: false, error: 'Usuario y contraseña requeridos' };
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('USUARIOS');

    if (!sheet) {
      return { success: false, error: 'No existe la hoja USUARIOS' };
    }

    const data = sheet.getDataRange().getValues();
    console.log('Filas USUARIOS:', data.length);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUsername = (row[0] || '').toString().trim();
      const rowPassword = (row[1] || '').toString().trim();
      const rowRol = (row[2] || '').toString().trim().toLowerCase();
      const rowDependencia = (row[3] || '').toString().trim();
      const rowToken = (row[4] || '').toString().trim();

      if (rowUsername.toLowerCase() === username.toLowerCase().trim() &&
          rowPassword === password.trim()) {
        
        console.log('✅ LOGIN OK:', rowUsername, 'Rol:', rowRol, 'Dep:', rowDependencia);
        
        return {
          success: true,
          usuario: {
            username: rowUsername,
            rol: rowRol,
            dependencia: rowDependencia,
            token: rowToken || Utilities.getUuid()
          }
        };
      }
    }

    console.log('❌ LOGIN FALLIDO');
    return { success: false, error: 'Usuario o contraseña incorrectos' };

  } catch (error) {
    console.error('ERROR handleLogin:', error);
    return { success: false, error: 'Error del servidor: ' + error };
  }
}

// ============================================================
// VERIFICAR CLAVE DE DEPENDENCIA
// ============================================================
function handleVerifyDepKey(username, depKey, dependencia) {
  console.log('=== handleVerifyDepKey ===');
  console.log('username:', username);
  console.log('depKey:', depKey);
  console.log('dependencia:', dependencia);

  if (!username || !depKey || !dependencia) {
    console.log('ERROR: Datos incompletos');
    return { success: false, error: 'Datos incompletos' };
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('USUARIOS');
    
    if (!sheet) {
      console.log('ERROR: No existe hoja USUARIOS');
      return { success: false, error: 'No existe la hoja USUARIOS' };
    }

    const data = sheet.getDataRange().getValues();
    console.log('Buscando en USUARIOS, total filas:', data.length);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUsername = (row[0] || '').toString().trim();
      const rowDependencia = (row[3] || '').toString().trim().toUpperCase();
      const rowDepKey = (row[4] || '').toString().trim();
      
      console.log('Fila', i, '- User:', rowUsername, '| Dep:', rowDependencia, '| Clave:', rowDepKey);
      
      if (rowUsername.toLowerCase() === username.toLowerCase().trim()) {
        console.log('✅ Usuario encontrado en fila', i);
        
        const depName = dependencia.toString().trim().toUpperCase();
        if (rowDependencia !== depName) {
          console.log('❌ Dependencia no coincide. Esperada:', depName, '| Encontrada:', rowDependencia);
          return { success: false, error: 'La dependencia del usuario no coincide' };
        }
        
        if (!rowDepKey) {
          console.log('ERROR: Columna E (clave) está vacía para este usuario');
          return { success: false, error: 'El usuario no tiene clave de dependencia configurada en columna E' };
        }
        
        const depKeyStr = depKey.toString().trim();
        console.log('Comparando - Guardada:', '"' + rowDepKey + '"', '| Ingresada:', '"' + depKeyStr + '"');
        
        if (rowDepKey === depKeyStr) {
          console.log('✅ CLAVE CORRECTA');
          return { success: true, message: 'Clave verificada' };
        } else {
          console.log('❌ CLAVE INCORRECTA');
          return { success: false, error: 'Clave de dependencia incorrecta' };
        }
      }
    }

    console.log('❌ Usuario no encontrado en USUARIOS');
    return { success: false, error: 'Usuario no encontrado' };

  } catch (error) {
    console.error('ERROR handleVerifyDepKey:', error);
    return { success: false, error: 'Error del servidor: ' + error.toString() };
  }
}

// ============================================================
// GET DATA - CORREGIDO: GRADO en columna D (índice 4)
// ============================================================
function handleGetData(dependency) {
  console.log('=== handleGetData === dependency=' + dependency);
  
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    
    if (!sheet) {
      console.log('ERROR: Hoja no encontrada:', dependency);
      return { success: false, error: 'Dependencia no encontrada: ' + dependency };
    }

    const lastRow = sheet.getLastRow();
    console.log('Última fila:', lastRow);
    
    if (lastRow <= 1) {
      console.log('Solo hay encabezados, sin datos');
      return { success: true, data: [] };
    }

    // Leer columnas A-H (8 columnas)
    const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const result = [];

    values.forEach(function (row, i) {
      const grado = row[CONFIG.COL_GRADO - 1];      // índice 3 = columna D
      const nombres = row[CONFIG.COL_NOMBRES - 1];  // índice 4 = columna E
      const turno = row[CONFIG.COL_TURNO - 1];      // índice 5 = columna F
      const novedad = row[CONFIG.COL_NOVEDAD - 1];  // índice 6 = columna G
      const observacion = row[CONFIG.COL_OBS - 1];  // índice 7 = columna H
      
      console.log('Fila', i+2, '- Grado:', grado, '| Nombre:', nombres, '| Turno:', turno, '| Novedad:', novedad);
      
      result.push({
        row: i + 2,
        grado: grado || '',
        nombres: nombres || '',
        turno: (turno || '').toString().trim().toUpperCase(),
        novedad: novedad || '',
        observacion: observacion || ''
      });
    });

    console.log('Total registros devueltos:', result.length);
    return { success: true, data: result };
    
  } catch (error) {
    console.error('Error getData:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// SAVE DATA
// ============================================================
function handleSaveData(dependency, params) {
  const row = parseInt(params.row);
  const novedad = params.novedad || '';
  const observation = params.observation || '';

  console.log('=== handleSaveData === dep=' + dependency + ' row=' + row);

  if (!row || isNaN(row) || row < 2) {
    return { success: false, error: 'Fila inválida: ' + params.row };
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    if (!sheet) return { success: false, error: 'Dependencia no encontrada: ' + dependency };

    const lastRow = sheet.getLastRow();
    if (row > lastRow) return { success: false, error: 'Fila ' + row + ' no existe. Última: ' + lastRow };

    sheet.getRange(row, CONFIG.COL_NOVEDAD).setValue(novedad);
    sheet.getRange(row, CONFIG.COL_OBS).setValue(observation);

    console.log('✅ Guardado fila', row);
    return { success: true, message: 'Datos guardados', row: row, novedad: novedad, observation: observation };
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
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    if (!sheet) return { success: false, error: 'Dependencia no encontrada' };
    const lastRow = sheet.getLastRow();
    return { success: true, total: lastRow > 1 ? lastRow - 1 : 0 };
  } catch (error) {
    console.error('Error:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GET SUMMARY
// ============================================================
function handleGetSummary(dependency, turno) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(dependency);
    if (!sheet) return { success: false, error: 'Dependencia no encontrada' };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, summary: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    let filteredData = values;
    if (turno) {
      filteredData = values.filter(function (row) {
        return (row[CONFIG.COL_TURNO - 1] || '').toString().trim().toUpperCase() === turno.toUpperCase();
      });
    }

    const total = filteredData.length;
    const conNovedad = filteredData.filter(function (r) {
      return (r[CONFIG.COL_NOVEDAD - 1] || '').toString().trim() !== '';
    }).length;
    const sinNovedad = total - conNovedad;

    const novedadesCount = {};
    filteredData.forEach(function (row) {
      const nov = (row[CONFIG.COL_NOVEDAD - 1] || '').toString().trim();
      if (nov) novedadesCount[nov] = (novedadesCount[nov] || 0) + 1;
    });

    const summary = [];
    summary.push({ section: 'header', label: 'RESUMEN ' + dependency + (turno ? ' - TURNO ' + turno : ''), value: '' });
    summary.push({ section: 'normal', label: 'Total Funcionarios', value: total });
    summary.push({ section: 'normal', label: 'Con Novedad', value: conNovedad });
    summary.push({ section: 'normal', label: 'Sin Novedad', value: sinNovedad });

    if (Object.keys(novedadesCount).length > 0) {
      summary.push({ section: 'novedad-header', label: 'DETALLE DE NOVEDADES', value: '' });
      for (const [novedad, count] of Object.entries(novedadesCount)) {
        summary.push({ section: 'novedad-item', label: novedad, value: count });
      }
    }

    return { success: true, summary: summary };
  } catch (error) {
    console.error('Error getSummary:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GET CELLS
// ============================================================
function handleGetCells(sheetName, cellsParam) {
  if (!sheetName || !cellsParam) {
    return { success: false, error: 'Parámetros sheet y cells requeridos' };
  }
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Hoja no encontrada: ' + sheetName };

    const celdas = cellsParam.split(',').map(function(c) { return c.trim(); });
    const valores = {};
    celdas.forEach(function(celda) {
      if (!celda) return;
      try { valores[celda] = sheet.getRange(celda).getValue(); } catch(e) { valores[celda] = ''; }
    });
    return { success: true, values: valores };
  } catch (error) {
    console.error('Error getCells:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GET AUDITORIA
// ============================================================
function handleGetAuditoria(dependency) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('AUDITORIA');
    if (!sheet) return { success: true, data: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = [];

    values.forEach(function(row) {
      const rowDep = (row[4] || '').toString().trim();
      if (dependency && rowDep && rowDep !== dependency) return;
      result.push({
        fechaHora: row[0] ? new Date(row[0]).toISOString() : '',
        usuario: (row[1] || '').toString().trim(),
        accion: (row[2] || '').toString().trim(),
        detalle: (row[3] || '').toString().trim()
      });
    });

    return { success: true, data: result };
  } catch (error) {
    console.error('Error getAuditoria:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// GET USUARIOS
// ============================================================
function handleGetUsuarios() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
    if (!sheet) return { success: true, data: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = [];

    values.forEach(function(row) {
      const user = (row[0] || '').toString().trim();
      if (!user) return;
      result.push({
        usuario: user,
        rol: (row[2] || '').toString().trim(),
        dependencia: (row[3] || '').toString().trim(),
        tieneClave: !!(row[4] || '').toString().trim()
      });
    });

    return { success: true, data: result };
  } catch (error) {
    console.error('Error getUsuarios:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// FUNCIONES DE ADMINISTRACIÓN
// ============================================================
function crearHojaDependencia(nombreDependencia, claveDependencia) {
  const ss = getSpreadsheet();
  if (ss.getSheetByName(nombreDependencia)) {
    return 'La hoja ' + nombreDependencia + ' ya existe';
  }

  const sheet = ss.insertSheet(nombreDependencia);
  const headers = ['ID', '', '', 'GRADO', 'APELLIDOS Y NOMBRES', 'TURNO', 'NOVEDAD', 'OBSERVACION'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#01592F').setFontColor('#FFFFFF');

  sheet.setColumnWidth(1, 50); 
  sheet.setColumnWidth(4, 80);   // D = GRADO
  sheet.setColumnWidth(5, 250);  // E = NOMBRES
  sheet.setColumnWidth(6, 60);   // F = TURNO
  sheet.setColumnWidth(7, 150);  // G = NOVEDAD
  sheet.setColumnWidth(8, 200);  // H = OBSERVACION

  const novedades = [
    'S/N','OFICINA','VACACIONES','SERVICIO','ESTADO DE GRAVIDEZ',
    'EXCUSA TOTAL','PERMISO','CITA MEDICA','LICENCIA DE MATERNIDAD',
    'EXCUSADO DEL SERVICIO','OTRA NOVEDAD','RETARDAD@','EXCUSA PARCIAL',
    'CURSO DE ASCENSO','COMISION'
  ];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(novedades, true)
    .setAllowInvalid(false)
    .setHelpText('Seleccione una novedad válida')
    .build();

  sheet.getRange('G2:G1000').setDataValidation(rule);
  sheet.setFrozenRows(1);

  console.log('✅ Hoja creada:', nombreDependencia);
  return 'Hoja ' + nombreDependencia + ' creada exitosamente';
}

function crearHojaUsuarios() {
  const ss = getSpreadsheet();
  if (ss.getSheetByName(CONFIG.USERS_SHEET)) {
    return 'La hoja USUARIOS ya existe';
  }
  const sheet = ss.insertSheet(CONFIG.USERS_SHEET);
  const headers = ['USUARIO', 'CONTRASEÑA', 'ROL', 'DEPENDENCIA', 'CLAVE_DEP'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#01592F').setFontColor('#FFFFFF');
  [120, 120, 100, 120, 120].forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
  return 'Hoja USUARIOS creada';
}

function agregarUsuario(username, password, rol, dependencia, claveDep) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return 'Error: Hoja USUARIOS no existe.';
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 5)
    .setValues([[username, password, rol, dependencia, claveDep || '']]);
  return 'Usuario ' + username + ' agregado';
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚔 Policía Nacional')
    .addItem('➕ Crear Hoja Dependencia', 'mostrarDialogoCrearDependencia')
    .addItem('👤 Crear Hoja Usuarios', 'crearHojaUsuarios')
    .addItem('➕ Agregar Usuario', 'mostrarDialogoAgregarUsuario')
    .addSeparator()
    .addItem('📋 Ver Dependencias', 'listarDependencias')
    .addItem('🔍 Ver Usuarios y Claves', 'mostrarUsuariosYClaves')
    .addToUi();
}

function mostrarDialogoCrearDependencia() {
  const html = HtmlService.createHtmlOutput(`
    <style>body{font-family:Arial,sans-serif;padding:20px}
    input{width:100%;padding:8px;margin:5px 0;border:1px solid #ccc;border-radius:4px}
    button{background:#01592F;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer}</style>
    <h3>Crear Nueva Dependencia</h3>
    <input type="text" id="nombre" placeholder="Nombre (ej: ASJUR)">
    <button onclick="crear()">Crear</button>
    <div id="resultado"></div>
    <script>
      function crear(){
        const n=document.getElementById('nombre').value.trim();
        if(!n){alert('Ingrese nombre');return;}
        google.script.run.withSuccessHandler(function(r){
          document.getElementById('resultado').innerHTML='<p style="color:green">'+r+'</p>';
        }).crearHojaDependencia(n);
      }
    <\/script>
  `).setWidth(350).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Crear Dependencia');
}

function mostrarDialogoAgregarUsuario() {
  const html = HtmlService.createHtmlOutput(`
    <style>body{font-family:Arial,sans-serif;padding:20px}
    input,select{width:100%;padding:8px;margin:5px 0;border:1px solid #ccc;border-radius:4px}
    button{background:#01592F;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer}</style>
    <h3>Agregar Usuario</h3>
    <input type="text" id="username" placeholder="Usuario">
    <input type="password" id="password" placeholder="Contraseña">
    <select id="rol"><option value="operador">Operador</option><option value="administrador">Administrador</option></select>
    <input type="text" id="dependencia" placeholder="Dependencia (ej: ASJUR)">
    <input type="text" id="claveDep" placeholder="Clave de Dependencia (col E)">
    <button onclick="agregar()">Agregar</button>
    <div id="resultado"></div>
    <script>
      function agregar(){
        const u=document.getElementById('username').value.trim();
        const p=document.getElementById('password').value.trim();
        const r=document.getElementById('rol').value;
        const d=document.getElementById('dependencia').value.trim();
        const c=document.getElementById('claveDep').value.trim();
        if(!u||!p||!d){alert('Complete usuario, contraseña y dependencia');return;}
        google.script.run.withSuccessHandler(function(res){
          document.getElementById('resultado').innerHTML='<p style="color:green">'+res+'</p>';
        }).agregarUsuario(u,p,r,d,c);
      }
    <\/script>
  `).setWidth(350).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Agregar Usuario');
}

function mostrarUsuariosYClaves() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('USUARIOS');
    if (!sheet) { SpreadsheetApp.getUi().alert('No existe hoja USUARIOS'); return; }
    
    const data = sheet.getDataRange().getValues();
    let msg = 'USUARIOS Y CLAVES DE DEPENDENCIA (Col E):\\n\\n';
    for (let i = 1; i < data.length; i++) {
      const user = (data[i][0] || '').toString().trim();
      const dep = (data[i][3] || '').toString().trim();
      const clave = (data[i][4] || '').toString().trim();
      if (user) {
        msg += user + ' | Dep: ' + dep + ' | Clave: ' + (clave || '⚠️ SIN CLAVE') + '\\n';
      }
    }
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.toString());
  }
}

function listarDependencias() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const nombres = sheets.map(function (s) { return s.getName(); })
                        .filter(function (n) { return n !== CONFIG.USERS_SHEET; });
  SpreadsheetApp.getUi().alert('Dependencias: ' + nombres.join(', '));
}

function test() {
  console.log('Test OK');
  console.log('Hojas:', getSpreadsheet().getSheets().map(function (s) { return s.getName(); }));
}