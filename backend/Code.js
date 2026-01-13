
//BACKEND DE FUNCIONES PRINCIPALES/LOGIN/REGISTRO

/**
 * Función principal para servir la aplicación web
 */
function doGet() {
  const title = 'PASTAS';
  const faviconUrl = 'https://alimentosdoria.com/wp-content/uploads/2023/01/logo-doria.png';

  return HtmlService.createTemplateFromFile('frontend/index')
    .evaluate()
    .setTitle(title)
    .setFaviconUrl(faviconUrl)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función para incluir archivos HTML (CSS y JavaScript)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtiene la lista de líderes desde la hoja "Lideres"
 * Formato esperado: Columna A = Nombre, Columna B = Cédula
 */
function getLeaders() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      throw new Error(`La hoja "${SHEETS.LIDERES}" no existe`);
    }

    const data = sheet.getDataRange().getValues();
    const leaders = [];

    // Saltar la primera fila si contiene encabezados
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Solo columna A
        leaders.push({
          info: row[0].toString().trim()
        });
      }
    }

    return leaders;
  } catch (error) {
    console.error('Error al obtener líderes:', error);
    throw new Error('No se pudieron cargar los líderes: ' + error.message);
  }
}

/**
 * Obtiene la información del líder desde el string (formato: "Nombre - Cédula")
 */
function getLeaderInfoFromString(leaderString) {
  try {
    console.log('🔍 Buscando información del líder:', leaderString);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      console.error('❌ No se encontró la hoja de líderes');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      console.warn('⚠️ Hoja de líderes vacía o solo tiene encabezados');
      return null;
    }

    // Parsear el string del líder para obtener la cédula
    let cedulaBuscada = '';
    let nombreBuscado = '';

    // Intentar diferentes formatos
    if (leaderString.includes(' - ')) {
      const parts = leaderString.split(' - ');
      if (parts.length >= 2) {
        nombreBuscado = parts[0].trim();
        cedulaBuscada = parts[1].trim();
      }
    } else {
      // Si no viene en formato esperado, usar el string completo como cédula
      cedulaBuscada = leaderString.trim();
    }

    console.log(`📝 Búsqueda - Cédula: "${cedulaBuscada}", Nombre: "${nombreBuscado}"`);

    // Buscar en la hoja de líderes 
    // Asumiendo: columna A = nombre, columna B = cédula, columna E = email
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Verificar que la fila tenga datos
      if (!row[0] && !row[1]) continue;

      const nombreSheet = String(row[0] || '').trim();
      const cedulaSheet = String(row[1] || '').trim();
      const emailSheet = row[4] ? String(row[4]).trim() : '';

      console.log(`📋 Fila ${i}: "${nombreSheet}" | "${cedulaSheet}" | "${emailSheet}"`);

      // Buscar por cédula (más confiable)
      if (cedulaSheet && cedulaSheet === cedulaBuscada) {
        console.log(`✅ Líder encontrado por cédula: ${nombreSheet}`);
        return {
          nombre: nombreSheet,
          cedula: cedulaSheet,
          email: emailSheet
        };
      }

      // Buscar por nombre si la cédula no coincide
      if (nombreBuscado && nombreSheet && nombreSheet.includes(nombreBuscado)) {
        console.log(`✅ Líder encontrado por nombre: ${nombreSheet}`);
        return {
          nombre: nombreSheet,
          cedula: cedulaSheet,
          email: emailSheet
        };
      }
    }

    console.warn('❌ Líder no encontrado en la hoja para:', leaderString);

    // Log de las primeras filas para debugging
    console.log('📊 Primeras filas de líderes:');
    for (let i = 1; i < Math.min(5, data.length); i++) {
      const row = data[i];
      console.log(`Fila ${i}: ${String(row[0])} | ${String(row[1])} | ${String(row[4])}`);
    }

    return null;

  } catch (error) {
    console.error('💥 Error crítico al obtener información del líder:', error);
    return null;
  }
}

// Conversor de fechas sin desfase UTC
function parseLocalDate(dateString) {
  if (!dateString) return new Date();

  const [datePart, timePart] = dateString.trim().split(' ');
  const [year, month, day] = datePart.split('-').map(Number);

  let hour = 0, minute = 0;
  if (timePart) {
    [hour, minute] = timePart.split(':').map(Number);
  }

  return new Date(year, month - 1, day, hour, minute);
}

//LOGIN

function validarCedula(cedula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LIDERES");
  const data = sheet.getDataRange().getValues();

  // Recorre desde la fila 2 (fila 1 son encabezados)
  for (let i = 1; i < data.length; i++) {
    const cedulaSheet = data[i][1];
    const rolSheet = data[i][2];
    const procesoSheet = data[i][3];
    const correoSheet = data[i][4];
    const empresaSheet = data[i][5];

    if (String(cedulaSheet) === String(cedula)) {
      return { success: true, rol: rolSheet, proceso_user: procesoSheet, correo: correoSheet, empresa: empresaSheet };
    }
  }

  return { success: false };
}

function getNombreByCedula(cedula) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      throw new Error(`La hoja "${SHEETS.LIDERES}" no existe`);
    }

    const data = sheet.getDataRange().getValues();

    // Buscar la cédula en la columna B (índice 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const cedulaSheet = String(row[1]).trim();

      if (cedulaSheet === String(cedula).trim()) {
        // Retornar el nombre de la columna A (índice 0)
        return {
          success: true,
          nombre: row[0] ? row[0].toString().trim() : 'Usuario'
        };
      }
    }

    return { success: false, nombre: '' };
  } catch (error) {
    console.error('Error al obtener nombre:', error);
    return { success: false, nombre: '' };
  }
}

/**
 * Registra un nuevo usuario en la hoja de líderes
 * Columnas: B=Cédula, C=Rol, D=Proceso, E=Correo, F=Empresa, G=Nombres, H=Apellidos
 */
function registrarUsuario(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      throw new Error('La hoja LIDERES no existe');
    }

    // Obtener datos de la columna B (cédulas) para verificar duplicados y encontrar última fila
    const lastRowSheet = sheet.getLastRow();
    const colB = sheet.getRange(1, 2, lastRowSheet, 1).getValues(); // Columna B completa

    // <CHANGE> Buscar la última fila con datos reales en columna B
    let nextRow = 1;
    for (let i = 0; i < colB.length; i++) {
      const cedulaExistente = String(colB[i][0] || '').trim();

      // Verificar si la cédula ya existe
      if (cedulaExistente === formData.cedula) {
        return {
          success: false,
          message: 'Esta cédula ya está registrada en el sistema'
        };
      }

      // Si la celda tiene datos, actualizar nextRow
      if (cedulaExistente !== '') {
        nextRow = i + 2; // +2 porque i es base 0 y queremos la siguiente fila
      }
    }

    // Convertir todos los campos a mayúsculas excepto el correo
    const rowData = [
      '',                                    // Columna A - Vacía
      formData.cedula.toUpperCase(),         // Columna B - Cédula
      'USUARIO',                             // Columna C - Rol
      formData.proceso.toUpperCase(),        // Columna D - Proceso
      formData.correo,                       // Columna E - Correo (sin cambios)
      formData.empresa.toUpperCase(),        // Columna F - Empresa
      formData.nombres.toUpperCase(),        // Columna G - Nombres
      formData.apellidos.toUpperCase()       // Columna H - Apellidos
    ];

    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    console.log('Usuario registrado en fila: ' + nextRow + ' - Cédula: ' + formData.cedula);

    return {
      success: true,
      message: 'Usuario registrado exitosamente'
    };

  } catch (error) {
    console.error('Error al registrar usuario:', error);
    return {
      success: false,
      message: 'Error al registrar: ' + error.message
    };
  }
}
