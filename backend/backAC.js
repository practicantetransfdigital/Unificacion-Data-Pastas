
//BACKEND DE LAS ANORMALIDADES CRITICAS POR MAQUINAS

function submitMaquinasReport(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.REPORTES_MAQUINAS);
      const headers = [
        'Fecha de Registro',
        'Mecanico Responsable',
        'Proceso',
        'AreaProceso',
        'Subsistema',
        'Anormalidad',
        'AreaResponsable',
        'Estado',
        'ID Reporte',
        'Criticidad' // <CHANGE> Agregada columna Criticidad
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    const reportId = 'MAQ-' + new Date().getTime();
    const fechaRegistro = parseLocalDate(formData.fecha);

    const rowData = [
      fechaRegistro,
      formData.mecanicoResponsable,
      formData.proceso,
      formData.areaProceso,
      formData.subsistema,
      formData.anormalidad,
      formData.areaResponsable,
      'Abierto',
      reportId,
      formData.criticidad || 'Media' // <CHANGE> Agregado campo criticidad
    ];

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(nextRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('Reporte de máquina guardado exitosamente con ID:', reportId);

    return {
      success: true,
      reportId: reportId,
      message: 'Reporte de máquina guardado exitosamente'
    };

  } catch (error) {
    console.error('Error al guardar reporte de máquina:', error);
    return {
      success: false,
      message: 'No se pudo guardar el reporte: ' + error.message
    };
  }
}

/**
 * Obtiene todos los reportes de máquinas desde la hoja de cálculo
 */
function getMaquinasReports() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const reports = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const reportId = row[8] || 'MAQ-' + new Date(row[0]).getTime();

      reports.push({
        id: reportId,
        fechaRegistro: row[0]
          ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : "",
        mecanicoResponsable: row[1] || "",
        proceso: row[2] || "",
        areaProceso: row[3] || "",
        subsistema: row[4] || "",
        anormalidad: row[5] || "",
        areaResponsable: row[6] || "",
        estado: row[7] || "Abierto",
        criticidad: row[9] || "Media" // <CHANGE> Agregado campo criticidad
      });
    }

    reports.sort((a, b) => new Date(b.fechaRegistro) - new Date(a.fechaRegistro));

    console.log("Reportes de máquinas obtenidos: " + reports.length);

    return reports;

  } catch (error) {
    Logger.log("Error al obtener reportes de máquinas: " + error);
    return [];
  }
}

/**
 * Actualiza el estado de un reporte de máquina
 */
function updateMaquinaStatus(reportId, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);

    if (!sheet) {
      throw new Error('La hoja de reportes de máquinas no existe');
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const currentId = data[i][8] || 'MAQ-' + new Date(data[i][0]).getTime();

      if (currentId === reportId) {
        sheet.getRange(i + 1, 8).setValue(newStatus);

        if (!data[i][8]) {
          sheet.getRange(i + 1, 9).setValue(reportId);
        }

        return {
          success: true,
          message: 'Estado actualizado exitosamente'
        };
      }
    }

    throw new Error('Reporte no encontrado');

  } catch (error) {
    console.error('Error al actualizar estado de máquina:', error);
    throw new Error('No se pudo actualizar el estado: ' + error.message);
  }
}

/**
 * Agrega un comentario a un reporte de máquina
 */
function addMaquinasCommentToReport(reportId, comment, autor) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "Reportes_Maquinas_COMENTARIOS";
  let sheet = ss.getSheetByName(sheetName);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID Reporte", "Autor", "Comentario", "Fecha"]);
  }

  sheet.appendRow([
    reportId,
    autor,
    comment,
    new Date()
  ]);

  return { success: true, message: "Comentario agregado exitosamente" };
}

/**
 * Obtiene los comentarios de un reporte de máquina específico
 */
function getMaquinasCommentsForReport(reportId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Maquinas_COMENTARIOS");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === reportId) {
      comments.push({
        autor: row[1],
        comentario: row[2],
        fecha: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
  }
  return comments;
}

/**
 * Obtiene el conteo de comentarios para todos los reportes de máquinas
 */
function getMaquinasCommentsCountForReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Maquinas_COMENTARIOS");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const counts = {};

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (!counts[id]) counts[id] = 0;
    counts[id]++;
  }

  return counts;
}

// <CHANGE> Nueva función para actualizar criticidad
function updateMaquinaCriticidad(reportId, newCriticidad) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);

    if (!sheet) {
      throw new Error('La hoja de reportes de máquinas no existe');
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const currentId = data[i][8] || 'MAQ-' + new Date(data[i][0]).getTime();

      if (currentId === reportId) {
        // Actualizar criticidad en columna J (índice 9)
        sheet.getRange(i + 1, 10).setValue(newCriticidad);

        if (!data[i][8]) {
          sheet.getRange(i + 1, 9).setValue(reportId);
        }

        return {
          success: true,
          message: 'Criticidad actualizada exitosamente'
        };
      }
    }

    throw new Error('Reporte no encontrado');

  } catch (error) {
    console.error('Error al actualizar criticidad de máquina:', error);
    throw new Error('No se pudo actualizar la criticidad: ' + error.message);
  }
}