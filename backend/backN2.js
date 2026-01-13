
//BACKEND DE LOS REPORTES N2

/**
 * Guarda un reporte N2 en la hoja correspondiente
 */
function submitN2Report(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.REPORTES_N2);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.REPORTES_N2);
      const headers = [
        'Fecha de Registro', 'Líder Responsable', 'Proceso', 'ZonaProceso',
        'Anormalidad', 'Proceso Responsable', 'Fecha Prevista Solución',
        'Estado', 'ID Reporte', 'Nombre y Cédula', 'Fotos'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    let fotosLinks = [];

    const getDirectDriveLink = (fileUrl) => {
      const match = fileUrl.match(/\/d\/(.*?)\//);
      if (!match) return fileUrl;
      const fileId = match[1];
      return `https://drive.google.com/uc?export=view&id=${fileId}`;
    };

    // 📸 Guardar fotos en Drive - IDÉNTICO A TARJETAS
    if (formData.fotos && formData.fotos.length > 0) {
      const folder = DriveApp.getFolderById(FOLDER_ID_N2);

      fotosLinks = formData.fotos.map((base64, i) => {
        try {
          const contentType = base64.split(';')[0].split(':')[1];
          const bytes = Utilities.base64Decode(base64.split(',')[1]);
          const blob = Utilities.newBlob(bytes, contentType, `foto_n2_${Date.now()}_${i + 1}.jpg`);
          const file = folder.createFile(blob);

          // ✅ HACER PÚBLICO EXPLÍCITAMENTE
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

          // ✅ CONVERSIÓN IDÉNTICA A TARJETAS
          return getDirectDriveLink(file.getUrl());

        } catch (error) {
          console.error(`Error guardando foto ${i + 1}:`, error);
          return null;
        }
      }).filter(link => link !== null);
    }

    // Generar ID único
    const reportId = 'N2-' + new Date().getTime();

    // Preparar datos
    const fechaRegistro = parseLocalDate(formData.fecha);
    const fechaSolucion = parseLocalDate(formData.fechaSolucion);

    const rowData = [
      fechaRegistro,
      formData.liderResponsable,
      formData.proceso,
      formData.zonaProceso,
      formData.anormalidad,
      formData.procesoResponsable,
      fechaSolucion,
      'Pendiente',
      reportId,
      formData.nombreCedula,
      JSON.stringify(fotosLinks) // ✅ Guardar solo URLs como tarjetas
    ];

    // Insertar en hoja
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    // Formatear
    sheet.getRange(nextRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(nextRow, 7).setNumberFormat('dd/mm/yyyy');

    console.log('📧 Intentando enviar correo de notificación...');
    const leaderInfo = getLeaderInfoFromString(formData.liderResponsable);

    if (leaderInfo && leaderInfo.email) {
      const emailEnviado = sendEmailToLeader(leaderInfo, formData, reportId);
      if (emailEnviado) {
        console.log('✅ Notificación por correo enviada exitosamente');
      } else {
        console.log('⚠️ No se pudo enviar la notificación por correo');
      }
    } else {
      console.warn('⚠️ No se pudo obtener información del líder para enviar correo');
    }

    return {
      success: true,
      reportId,
      message: 'Reporte N2 guardado exitosamente',
      fotos: fotosLinks
    };

  } catch (error) {
    console.error('Error al guardar reporte N2:', error);
    throw new Error('No se pudo guardar el reporte: ' + error.message);
  }
}

/**
 * Obtiene todos los reportes N2 desde la hoja de cálculo
 */
function getN2Reports() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const reports = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      // CONVERTIR FOTOS DE DRIVE A BASE64
      let fotosBase64 = [];
      if (row[10]) {
        try {
          const urls = JSON.parse(row[10]);
          fotosBase64 = urls.map(url => {
            const idMatch = url.match(/id=([a-zA-Z0-9_-]+)/);
            if (!idMatch) return '';

            try {
              const file = DriveApp.getFileById(idMatch[1]);
              const blob = file.getBlob();
              return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
            } catch (e) {
              console.log('Error convirtiendo URL a Base64:', e);
              return '';
            }
          }).filter(base64 => base64 !== '');
        } catch (e) {
          console.log('Error parseando fotos para fila ' + i + ': ' + e);
          fotosBase64 = [];
        }
      }

      // Obtener responsable y fecha de cierre N2 (tus columnas específicas)
      let responsableCierreN2 = row[11] || ""; // Columna L = ResponsableCierreN2 (índice 11)
      let fechaCierreN2 = "";
      if (row[12]) { // Columna M = FechaCierreN2 (índice 12)
        fechaCierreN2 = Utilities.formatDate(
          new Date(row[12]),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm:ss"
        );
      }

      reports.push({
        fechaRegistro: row[0]
          ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : "",
        liderResponsable: row[1] || "",
        proceso: row[2] || "",
        zonaProceso: row[3] || "",
        anormalidad: row[4] || "",
        procesoResponsable: row[5] || "",
        fechaSolucion: row[6]
          ? Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : "",
        estado: row[7] || "Pendiente",
        id: row[8] || "",
        nombreCedula: row[9] || "",
        fotos: fotosBase64,
        // Nuevos campos con los nombres específicos de tu hoja
        responsableCierreN2: responsableCierreN2,
        fechaCierreN2: fechaCierreN2
      });
    }

    // Ordenar por fecha descendente
    reports.sort((a, b) => new Date(b.fechaRegistro) - new Date(a.fechaRegistro));

    console.log("✅ Reportes N2 obtenidos: " + reports.length);
    return reports;

  } catch (error) {
    Logger.log("❌ Error al obtener reportes N2: " + error);
    return [];
  }
}

/**
 * Actualiza el estado de un reporte N2
 * Cuando se cambia a "Completado", registra automáticamente responsable (nombre) y fecha
 */
function updateReportStatus(reportId, newStatus, nombreUsuario) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);

    if (!sheet) {
      throw new Error('La hoja de reportes N2 no existe');
    }

    const data = sheet.getDataRange().getValues();
    let reportRow = -1;

    // Buscar el reporte por ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][8] === reportId) {
        reportRow = i + 1;

        // Guardar el estado anterior
        const oldStatus = data[i][7] || '';

        // Actualizar el estado
        sheet.getRange(reportRow, 8).setValue(newStatus);

        if (newStatus === 'Completado' && oldStatus !== 'Completado') {
          const fechaCompletado = new Date();
          const responsableNombre = nombreUsuario || 'Usuario del sistema';

          // Guardar en las columnas específicas
          sheet.getRange(reportRow, 12).setValue(responsableNombre); // ResponsableCierreN2
          sheet.getRange(reportRow, 13).setValue(fechaCompletado); // FechaCierreN2

          // Formatear la fecha
          sheet.getRange(reportRow, 13).setNumberFormat('dd/mm/yyyy hh:mm');

          console.log(`✅ Reporte ${reportId} completado por: ${responsableNombre} a las ${fechaCompletado}`);
        }

        // Si se cambia de "Completado" a otro estado, limpiar los campos
        if (oldStatus === 'Completado' && newStatus !== 'Completado') {
          sheet.getRange(reportRow, 12).clearContent(); // Limpiar ResponsableCierreN2
          sheet.getRange(reportRow, 13).clearContent(); // Limpiar FechaCierreN2
          console.log(`⚠️ Reporte ${reportId} cambiado de Completado a ${newStatus}, campos limpiados`);
        }

        return {
          success: true,
          message: 'Estado actualizado exitosamente',
          changedToCompleted: (newStatus === 'Completado' && oldStatus !== 'Completado')
        };
      }
    }

    throw new Error('Reporte no encontrado');

  } catch (error) {
    console.error('Error al actualizar estado:', error);
    throw new Error('No se pudo actualizar el estado: ' + error.message);
  }
}

/**
 * Agrega un comentario a un reporte N2
 */
function addCommentToReport(reportId, comment, autor) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2 + "_COMENTARIOS") || ss.insertSheet(SHEETS.REPORTES_N2 + "_COMENTARIOS");

  // Si la hoja es nueva, crea encabezados
  if (sheet.getLastRow() === 0) {
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
 * Obtiene los comentarios de un reporte N2 específico
 */
function getCommentsForReport(reportId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2 + "_COMENTARIOS");
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
 * Cambia el responsable de un reporte
 */
function updateReportResponsible(reportId, newResponsible) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);
  if (!sheet) throw new Error("No existe la hoja N2");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === reportId) {
      sheet.getRange(i + 1, 2).setValue(newResponsible);
      return { success: true, message: "Responsable actualizado correctamente" };
    }
  }

  throw new Error("Reporte no encontrado");
}

function getCommentsCountForReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2 + "_COMENTARIOS");
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

function updateReportDate(reportId, newDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);            // usa mismo origen que updateReportResponsible
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);          // usa la misma constante/hoja
  if (!sheet) throw new Error("No existe la hoja N2");

  const data = sheet.getDataRange().getValues();


  const idCol = 8;     // columna donde está el ID (ej. H -> 7, H index 8 si antes lo usaste así)
  const fechaCol = 6;  // columna donde quieres escribir la fecha (ajusta si es necesario)

  Logger.log("📩 ID recibido desde frontend: %s", reportId);
  Logger.log("🗂 Ejemplo IDs (filas 2-6): %s", data.slice(1, 6).map(r => r[idCol]).join(", "));

  const normalizedTarget = String(reportId).trim();

  for (let i = 1; i < data.length; i++) {
    const cellId = String(data[i][idCol]).trim();
    Logger.log("Comparando fila %d: hojaId=%s target=%s", i + 1, cellId, normalizedTarget);

    if (cellId === normalizedTarget) {
      // intentar guardar como DATE (si newDate viene 'YYYY-MM-DD' lo convertimos)
      let valueToWrite = newDate;
      try {
        // Si newDate es string "YYYY-MM-DD", esto lo convierte a Date
        const parsed = new Date(newDate);
        if (!isNaN(parsed.getTime())) {
          parsed.setMinutes(parsed.getMinutes() + parsed.getTimezoneOffset());
          valueToWrite = parsed;
        }
      } catch (e) {
        // si falla, dejamos el string (setValue aceptará string también)
      }

      sheet.getRange(i + 1, fechaCol + 1).setValue(valueToWrite);
      Logger.log("✅ Fecha actualizada para ID %s en fila %d", reportId, i + 1);
      return { success: true, message: "Fecha actualizada correctamente" };
    }
  }

  // si no encontró, devolver info para debugging (no throw si prefieres manejarlo en frontend)
  Logger.log("❌ No se encontró el reporte con ID %s", reportId);
  throw new Error("No se encontró el reporte con ID " + reportId);
}