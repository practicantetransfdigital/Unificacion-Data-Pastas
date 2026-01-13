
//BACKEND DEL CONSOLIDADO DE REPORTES

// <CHANGE> Obtiene todos los reportes consolidados
function getConsolidadoReports() {
  try {
    const n2 = getN2Reports();
    const tarjetas = getTarjetasReports();
    const maquinas = getMaquinasReports();
    const ciclos = getCiclosMejora();

    return {
      n2: n2,
      tarjetas: tarjetas,
      maquinas: maquinas,
      ciclos: ciclos
    };

  } catch (error) {
    console.error('Error al obtener consolidado:', error);
    throw new Error('No se pudo cargar el consolidado: ' + error.message);
  }
}

function addTarjetaCommentToReport(tarjetaId, comment, autor) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "Reportes_Tarjetas_COMENTARIOS";
  let sheet = ss.getSheetByName(sheetName);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID Tarjeta", "Autor", "Comentario", "Fecha"]);
  }

  sheet.appendRow([
    tarjetaId,
    autor,
    comment,
    new Date()
  ]);

  return { success: true, message: "Comentario agregado exitosamente" };
}