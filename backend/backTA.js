
//BACKEND DE LAS TARJETAS DE ANORMALIDADES

function submitTarjetaReport(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);

    if (!sheet) throw new Error(`La hoja "${SHEETS.REPORTES_TARJETAS}" no existe`);

    let fotosLinks = [];

    // Función para convertir link de Drive a link directo
    const getDirectDriveLink = (fileUrl) => {
      const match = fileUrl.match(/\/d\/(.*?)\//);
      if (!match) return fileUrl;
      const fileId = match[1];
      return `https://drive.google.com/uc?export=view&id=${fileId}`;
    };

    // 📸 Guardar fotos en Drive si existen
    if (data.fotos && data.fotos.length > 0) {
      const folder = DriveApp.getFolderById(FOLDER_ID_TARJETAS);
      fotosLinks = data.fotos.map((base64, i) => {
        const contentType = base64.split(';')[0].split(':')[1];
        const bytes = Utilities.base64Decode(base64.split(',')[1]);
        const blob = Utilities.newBlob(bytes, contentType, `foto_${Date.now()}_${i + 1}.jpg`);
        const file = folder.createFile(blob);
        // Convertimos a link directo
        return getDirectDriveLink(file.getUrl());
      });
    }

    const totalTarjetas = sheet.getLastRow() - 1;
    const tarjetaId = `TAR-${String(totalTarjetas + 1).padStart(4, '0')}`;


    // CORRECCIÓN: Array con todas las columnas en el orden correcto
    const newRow = [
      data.zonaRiesgo || '',
      data.nombreCedula || '',
      data.ubicacion || '',
      data.prioridad || '',
      data.descripcionProblema || '',
      data.tipoRiesgo || '',
      data.problemaAsociado || '',
      data.sistemaGestion || '',
      data.responsableSolucion || '',
      data.generadaPor || '',
      data.fechaCreacionTarjeta || '',
      data.estado || 'Abierta',
      JSON.stringify(fotosLinks),
      '',
      '',
      data.requiereSAP || 'No',
      tarjetaId
    ];

    sheet.appendRow(newRow);

    const creadorEmail = getEmailByNombre(data.nombreCedula);
    const responsableEmail = RESPONSABLES_EMAILS[data.responsableSolucion] || RESPONSABLES_EMAILS["Por Asignar"];

    // Enviar correos
    if (creadorEmail) {
      sendEmailToCreador(creadorEmail, data, fotosLinks);
    }

    if (responsableEmail) {
      sendEmailToResponsable(responsableEmail, data, fotosLinks, creadorEmail);
    }

    return {
      success: true,
      tarjetaId: tarjetaId,
      message: 'Tarjeta de anormalidad registrada exitosamente',
      fotos: fotosLinks
    };
  } catch (error) {
    console.error('Error al guardar tarjeta:', error);
    return {
      success: false,
      message: 'Error al guardar la tarjeta: ' + error.message
    };
  }
}

function getTarjetasReports() {
  try {
    console.log('🔍 Iniciando getTarjetasReports...');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);
    if (!sheet) {
      console.log('❌ Hoja no encontrada:', SHEETS.REPORTES_TARJETAS);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    console.log('📊 Filas totales:', data.length);

    const tarjetas = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Verificar que row exista y tenga al menos 1 elemento
      if (!row || row.length === 0 || !row[0]) continue;

      // Convertir fotos de Drive a Base64
      let fotosUrls = [];
      if (row[12]) { // Columna 13: fotos (JSON con URLs)
        try {
          // Parsear JSON de URLs
          fotosUrls = JSON.parse(row[12]);

          // Asegurar que sea array
          if (!Array.isArray(fotosUrls)) {
            fotosUrls = [];
          }

          console.log(`Fila ${i + 1}: ${fotosUrls.length} URLs de fotos`);

        } catch (e) {
          console.error('Error parseando JSON de fotos (fila ' + i + '):', e);
          fotosUrls = [];
        }
      }

      // MANEJO SEGURO DEL ID - CRÍTICO
      let tarjetaId;
      try {
        tarjetaId = row[16];
        if (!tarjetaId) {
          // Si no hay ID en row[16], crear uno basado en fecha
          if (row[10]) {
            const fecha = new Date(row[10]);
            if (!isNaN(fecha.getTime())) {
              tarjetaId = 'TAR-' + fecha.getTime();
            } else {
              tarjetaId = 'TAR-' + new Date().getTime();
            }
          } else {
            tarjetaId = 'TAR-' + new Date().getTime();
          }
        }
      } catch (e) {
        console.error('Error generando ID para fila', i, ':', e);
        tarjetaId = 'TAR-' + new Date().getTime(); // Fallback
      }

      // Fecha de creación segura
      let fechaCreacion = "";
      try {
        if (row[10]) {
          const fecha = new Date(row[10]);
          if (!isNaN(fecha.getTime())) {
            fechaCreacion = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
        }
      } catch (e) {
        console.error('Error procesando fecha creación fila', i, ':', e);
      }

      // Fecha de cierre segura (si existe)
      let fechaCierreTarjeta = "";
      if (row.length > 24 && row[24]) {
        try {
          const fechaCierre = new Date(row[24]);
          if (!isNaN(fechaCierre.getTime())) {
            fechaCierreTarjeta = Utilities.formatDate(fechaCierre, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
        } catch (e) {
          console.error('Error procesando fecha cierre fila', i, ':', e);
        }
      }

      tarjetas.push({
        rowIndex: i + 1,
        id: tarjetaId,
        zonaRiesgo: row[0] || "",
        nombreCedula: row[1] || "",
        ubicacion: row[2] || "",
        prioridad: row[3] || "",
        descripcionProblema: row[4] || "",
        tipoRiesgo: row[5] || "",
        problemaAsociado: row[6] || "",
        sistemaGestion: row[7] || "",
        responsableSolucion: row[8] || "",
        responsableNombreVisualizarReporte: row[18] || "",
        generadaPor: row[9] || "",
        fechaCreacion: fechaCreacion,
        estado: row[11] || "Abierta",
        fotos: fotosUrls,
        comentarioCierre: row[13] || "",
        responsableCierre: row[14] || "",
        requiereSAP: row[15] || "No",
        fechaCierreTarjeta: fechaCierreTarjeta
      });
    }

    console.log('✅ Tarjetas procesadas:', tarjetas.length);

    // Ordenar por prioridad solo si hay tarjetas
    if (tarjetas.length > 0) {
      const prioridadOrden = { "Alta": 1, "Media": 2, "Baja": 3 };
      tarjetas.sort((a, b) => (prioridadOrden[a.prioridad] || 999) - (prioridadOrden[b.prioridad] || 999));
    }

    return tarjetas;

  } catch (error) {
    console.error("❌ Error CRÍTICO en getTarjetasReports:", error);
    console.error("❌ Stack trace:", error.stack);
    return [];
  }
}

/**
 * Cierra una tarjeta de anormalidad con un comentario
 */
function closeTarjetaReport(rowIndex, comentario, responsableCierre, fechaCierreTarjeta = null) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);

    if (!sheet) {
      throw new Error('La hoja de tarjetas no existe');
    }

    // Usar la fecha pasada como parámetro o la fecha actual
    const fechaCierre = fechaCierreTarjeta
      ? Utilities.formatDate(new Date(fechaCierreTarjeta), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    sheet.getRange(rowIndex, 12).setValue('Cerrada');
    sheet.getRange(rowIndex, 14).setValue(comentario);
    sheet.getRange(rowIndex, 15).setValue(responsableCierre);
    sheet.getRange(rowIndex, 25).setValue(fechaCierre);

    return {
      success: true,
      message: 'Tarjeta cerrada exitosamente'
    };

  } catch (error) {
    console.error('Error al cerrar tarjeta:', error);
    throw new Error('No se pudo cerrar la tarjeta: ' + error.message);
  }
}

function getTarjetasCommentsForReport(tarjetaId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Tarjetas_COMENTARIOS");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === tarjetaId) {
      comments.push({
        autor: row[1],
        comentario: row[2],
        fecha: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
  }
  return comments;
}

function getTarjetasCommentsCountForReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Tarjetas_COMENTARIOS");
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

/**
 * Actualiza el responsable de una tarjeta
 */
function updateTarjetaResponsible(rowIndex, newResponsible) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);

    if (!sheet) {
      throw new Error('La hoja de tarjetas no existe');
    }

    // Actualizar en la columna 9 (índice 9 = columna I = Responsable)
    sheet.getRange(rowIndex, 9).setValue(newResponsible);

    return {
      success: true,
      message: 'Responsable actualizado correctamente'
    };

  } catch (error) {
    console.error('Error al actualizar responsable de tarjeta:', error);
    throw new Error('No se pudo actualizar el responsable: ' + error.message);
  }
}