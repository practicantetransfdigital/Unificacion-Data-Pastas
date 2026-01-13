
// BACKEND DE LOS CICLOS DE MEJORA

function getNextCicloId() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('REPORTES_CICLOS');

    if (!sheet) {
      return 'CM-001';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return 'CM-001';
    }

    const totalCiclos = lastRow - 1;
    const nextNumber = totalCiclos + 1;
    return 'CM-' + String(nextNumber).padStart(3, '0');

  } catch (error) {
    console.error('Error obteniendo siguiente ID de ciclo:', error);
    return 'CM-001';
  }
}

function submitCicloMejora(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Reportes_Ciclos');

    if (!sheet) {
      sheet = ss.insertSheet('Reportes_Ciclos');
      const headers = [
        'ID Ciclo', 'Fecha Registro', 'Nombre Ciclo', 'Aviso Mantenimiento',
        'Proceso', 'Equipo/Máquina', 'Líder', 'Integrantes',
        'Tipo Foco Mejora', 'Datos Foco Mejora',
        'Defecto Principal',
        'Causas Medio Ambiente', 'Causas Mano de Obra', 'Causas Materiales',
        'Causas Tiempo', 'Causas Método', 'Causas Máquina',
        'Análisis 5 Por Qué', 'Plan de Acción 5W+2H', 'Estado', 'Creado Por'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, headers.length).setBackground('#0f307f');
      sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
    }

    const cicloId = formData.cicloId || getNextCicloId();
    const fechaRegistro = parseLocalDate(formData.fecha);

    const espina = formData.espinaPescado || {};
    const causasAmbiente = (espina.medioAmbiente || []).join(' | ');
    const causasMano = (espina.manoDeObra || []).join(' | ');
    const causasMateriales = (espina.materiales || []).join(' | ');
    const causasTiempo = (espina.tiempo || []).join(' | ');
    const causasMetodo = (espina.metodo || []).join(' | ');
    const causasMaquina = (espina.maquina || []).join(' | ');

    const analisis5PorquesStr = formData.analisis5Porques ?
      JSON.stringify(formData.analisis5Porques) : '';

    // 🔥 CORRECCIÓN: GUARDAR PLAN DE ACCIÓN LIMPIO
    const focoMejora = formData.focoMejora || {};
    const tipoFoco = focoMejora.tipo || '';
    const datosFocoStr = JSON.stringify(focoMejora);

    let planAccionStr = '';
    if (formData.planAccion && Array.isArray(formData.planAccion)) {
      // Limpiar el array antes de convertirlo a string
      const planAccionLimpio = formData.planAccion.map(accion => {
        return {
          cual: (accion.cual || '').replace(/["'\\$%]/g, ''),
          que: (accion.que || '').replace(/["'\\$%]/g, ''),
          donde: (accion.donde || '').replace(/["'\\$%]/g, ''),
          quien: (accion.quien || '').replace(/["'\\$%]/g, ''),
          como: (accion.como || '').replace(/["'\\$%]/g, ''),
          cuando: (accion.cuando || '').replace(/["'\\$%]/g, ''),
          cuanto: (accion.cuanto || '').replace(/["'\\$%]/g, '')
        };
      });

      planAccionStr = JSON.stringify(planAccionLimpio);
      console.log('[Backend] Plan acción guardado (limpio):', planAccionStr.substring(0, 200));
    }

    const rowData = [
      cicloId,
      fechaRegistro,
      formData.nombreCiclo || '',
      formData.avisoMantenimiento || '',
      formData.proceso || '',
      formData.equipoMaquina || '',
      formData.lider || '',
      formData.integrantes || '',
      tipoFoco,
      datosFocoStr,
      (formData.defecto || '').replace(/["'\\$%]/g, ''), // Limpiar defecto también
      causasAmbiente,
      causasMano,
      causasMateriales,
      causasTiempo,
      causasMetodo,
      causasMaquina,
      analisis5PorquesStr,
      planAccionStr,
      'Abierto',
      formData.creadoPor || ''
    ];

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(nextRow, 2).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('Ciclo de Mejora guardado exitosamente con ID:', cicloId);

    return {
      success: true,
      cicloId: cicloId,
      message: 'Ciclo de Mejora registrado exitosamente'
    };

  } catch (error) {
    console.error('Error al guardar Ciclo de Mejora:', error);
    return {
      success: false,
      message: 'Error al guardar: ' + error.message
    };
  }
}

// ========== FUNCIONES GESTIÓN DE CICLOS ==========

function getCiclosMejora() {
  try {
    console.log('[Backend] Iniciando getCiclosMejora...');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Reportes_Ciclos');

    if (!sheet) {
      sheet = ss.getSheetByName('CICLOS_MEJORA');
    }

    if (!sheet) {
      console.log('[Backend] ERROR: Ninguna hoja de ciclos encontrada');
      return [];
    }

    console.log('[Backend] Hoja encontrada:', sheet.getName());

    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('[Backend] Solo encabezados, sin datos');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    console.log('[Backend] Datos obtenidos, filas:', data.length);

    const ciclos = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Saltar filas completamente vacías
      if (!row[0] && !row[1] && !row[2]) continue;

      // CONVERTIR FECHA A STRING ISO
      let fechaStr = '';
      try {
        if (row[1] instanceof Date) {
          fechaStr = row[1].toISOString();
        } else if (row[1]) {
          fechaStr = new Date(row[1]).toISOString();
        }
      } catch (e) {
        fechaStr = '';
      }

      // 🔥 DIAGNÓSTICO: Ver el plan de acción crudo
      let planAccionStr = String(row[18] || '').trim();
      console.log(`[Backend] Fila ${i + 1} - Plan acción CRUDO (primeros 500 chars):`,
        planAccionStr.substring(0, 500));
      console.log(`[Backend] Fila ${i + 1} - Plan acción longitud:`, planAccionStr.length);

      let planAccionParseado = [];

      if (planAccionStr && planAccionStr !== '') {
        try {
          // 1. Limpiar caracteres básicos
          planAccionStr = planAccionStr
            .replace(/^[\x00-\x1F]+/, '')
            .trim();

          // 2. DIAGNÓSTICO: Ver el JSON exacto
          console.log(`[Backend] JSON antes de parsear (primeros 500 chars):`,
            planAccionStr.substring(0, 500));

          // 3. Intentar parsear directamente primero
          try {
            planAccionParseado = JSON.parse(planAccionStr);
            console.log(`[Backend] Parseado directo exitoso, ${planAccionParseado.length} acciones`);
          } catch (parseError1) {
            console.log(`[Backend] Primer intento falló: ${parseError1.message}`);

            // 4. Intentar con función de reparación
            try {
              planAccionParseado = repararJSONPlanAccion(planAccionStr);
              console.log(`[Backend] JSON reparado exitoso, ${planAccionParseado.length} acciones`);
            } catch (parseError2) {
              console.error(`[Backend] Segundo intento falló: ${parseError2.message}`);

              // 5. Último intento: extraer solo el array manualmente
              try {
                // Buscar el array completo entre [ y ]
                const inicio = planAccionStr.indexOf('[');
                const fin = planAccionStr.lastIndexOf(']');

                if (inicio !== -1 && fin !== -1 && fin > inicio) {
                  const jsonExtraido = planAccionStr.substring(inicio, fin + 1);
                  // Limpiar escapes de barras invertidas
                  const jsonLimpio = jsonExtraido.replace(/\\\\/g, '\\');
                  planAccionParseado = JSON.parse(jsonLimpio);
                  console.log(`[Backend] Extracción manual exitosa, ${planAccionParseado.length} acciones`);
                } else {
                  throw new Error('No se encontró array en el string');
                }
              } catch (parseError3) {
                console.error(`[Backend] Tercer intento falló: ${parseError3.message}`);
                planAccionParseado = [];
              }
            }
          }
        } catch (error) {
          console.error(`❌ Error procesando plan de acción en fila ${i + 1}:`, error);
          planAccionParseado = [];
        }

      }

      const ciclo = {
        id: String(row[0] || '').trim(),
        fecha: fechaStr,
        nombre: String(row[2] || '').trim(),
        aviso: String(row[3] || '').trim(),
        proceso: String(row[4] || '').trim(),
        equipo: String(row[5] || '').trim(),
        lider: String(row[6] || '').trim(),
        integrantes: String(row[7] || '').trim(),
        tipoFoco: String(row[8] || '').trim(),
        datosFoco: String(row[9] || '').trim(),
        defecto: String(row[10] || '').trim(),
        causasAmbiente: String(row[11] || '').trim(),
        causasMano: String(row[12] || '').trim(),
        causasMateriales: String(row[13] || '').trim(),
        causasTiempo: String(row[14] || '').trim(),
        causasMetodo: String(row[15] || '').trim(),
        causasMaquina: String(row[16] || '').trim(),
        analisis5Porques: String(row[17] || '').trim(),
        planAccion: planAccionParseado,
        estado: String(row[19] || 'Abierto').trim(),
        creadoPor: String(row[20] || '').trim()
      };

      ciclos.push(ciclo);
    }

    console.log('[Backend] Ciclos procesados:', ciclos.length);

    // Mostrar detalles del primer ciclo para diagnóstico
    if (ciclos.length > 0) {
      console.log('[Backend] Primer ciclo ID:', ciclos[0].id);
      console.log('[Backend] Primer ciclo - planAccion tipo:', typeof ciclos[0].planAccion);
      console.log('[Backend] Primer ciclo - planAccion es array?', Array.isArray(ciclos[0].planAccion));
      if (Array.isArray(ciclos[0].planAccion)) {
        console.log('[Backend] Primer ciclo - acciones:', ciclos[0].planAccion.length);
        if (ciclos[0].planAccion.length > 0) {
          console.log('[Backend] Primera acción:', JSON.stringify(ciclos[0].planAccion[0]).substring(0, 200));
        }
      }
    }

    // Ordenar por fecha descendente
    if (ciclos.length > 1) {
      ciclos.sort((a, b) => {
        try {
          const dateA = a.fecha ? new Date(a.fecha).getTime() : 0;
          const dateB = b.fecha ? new Date(b.fecha).getTime() : 0;
          return dateB - dateA;
        } catch (e) {
          return 0;
        }
      });
    }

    return ciclos;

  } catch (error) {
    console.error('[Backend] ERROR en getCiclosMejora:', error);
    console.error('[Backend] Stack trace:', error.stack);
    return [];
  }
}

// Obtener historial de seguimiento de un ciclo 
function getHistorialCiclo(cicloId) {
  try {
    console.log('🔍 Buscando historial para ciclo:', cicloId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('CICLOS_HISTORIAL');

    // Si no existe la hoja, crear una nueva
    if (!sheet) {
      console.log('📝 Creando nueva hoja de historial');
      sheet = ss.insertSheet('CICLOS_HISTORIAL');
      sheet.appendRow(['ID Ciclo', 'Fecha', 'Estado', 'Comentario', 'Autor']);
      sheet.getRange(1, 1, 1, 5).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
      return []; // Retornar vacío porque es nueva
    }

    const data = sheet.getDataRange().getValues();
    console.log('📊 Datos en hoja de historial:', data.length, 'filas');

    if (data.length <= 1) {
      console.log('📭 Hoja de historial vacía o solo encabezados');
      return [];
    }

    const historial = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();

      console.log(`Fila ${i}: ID="${rowCicloId}" buscando="${cicloId}"`);

      if (rowCicloId === cicloId) {
        const registro = {
          cicloId: rowCicloId,
          fecha: row[1] ? row[1].toISOString() : new Date().toISOString(),
          estado: String(row[2] || ''),
          comentario: String(row[3] || ''),
          autor: String(row[4] || 'Sistema')
        };

        console.log('✅ Registro encontrado:', registro);
        historial.push(registro);
      }
    }

    console.log('📋 Total registros encontrados:', historial.length);
    return historial;

  } catch (error) {
    console.error('💥 Error crítico en getHistorialCiclo:', error);
    return [];
  }
}

// Guardar seguimiento de ciclo - VERSIÓN MEJORADA
function guardarSeguimientoCiclo(seguimiento) {
  try {
    console.log('💾 Guardando seguimiento para ciclo:', seguimiento.cicloId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Guardar en historial
    let historialSheet = ss.getSheetByName('CICLOS_HISTORIAL');
    if (!historialSheet) {
      console.log('📝 Creando nueva hoja de historial');
      historialSheet = ss.insertSheet('CICLOS_HISTORIAL');
      historialSheet.appendRow(['ID Ciclo', 'Fecha', 'Estado', 'Comentario', 'Autor']);
      historialSheet.getRange(1, 1, 1, 5).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
    }

    const fechaActual = new Date();

    // Agregar registro al historial
    historialSheet.appendRow([
      seguimiento.cicloId,
      fechaActual,
      seguimiento.estado,
      seguimiento.comentario,
      seguimiento.autor || 'Usuario desconocido'
    ]);

    // Formatear la fecha en la hoja
    const lastRow = historialSheet.getLastRow();
    historialSheet.getRange(lastRow, 2).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('✅ Seguimiento guardado en historial, fila:', lastRow);

    // 2. Actualizar estado en hoja de ciclos
    let ciclosSheet = ss.getSheetByName('Reportes_Ciclos');
    if (!ciclosSheet) {
      ciclosSheet = ss.getSheetByName('CICLOS_MEJORA'); // Buscar con otro nombre
    }

    if (ciclosSheet) {
      const data = ciclosSheet.getDataRange().getValues();
      let encontrado = false;

      for (let i = 1; i < data.length; i++) {
        const rowId = String(data[i][0] || '').trim();
        if (rowId === seguimiento.cicloId) {
          // Columna 20 es el estado (índice 19)
          ciclosSheet.getRange(i + 1, 20).setValue(seguimiento.estado);
          console.log('✅ Estado actualizado en hoja de ciclos, fila:', i + 1);
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        console.warn('⚠️ Ciclo no encontrado en hoja principal:', seguimiento.cicloId);
      }
    } else {
      console.warn('⚠️ Hoja de ciclos no encontrada');
    }

    return {
      success: true,
      message: 'Seguimiento guardado correctamente',
      detalles: {
        historialRow: lastRow,
        fecha: fechaActual.toISOString()
      }
    };

  } catch (error) {
    console.error('❌ Error en guardarSeguimientoCiclo:', error);
    return {
      success: false,
      message: 'Error al guardar seguimiento: ' + error.toString()
    };
  }
}

function getAccionesCausa(cicloId, idCausa) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccionesCausa');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const acciones = [];
    const idCausaNum = parseInt(idCausa);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = row[0] ? row[0].toString().trim() : '';
      const rowIdCausa = row[1] !== undefined ? parseInt(row[1]) : null;

      if (rowCicloId === cicloId.toString() && rowIdCausa === idCausaNum) {
        const accion = {
          que: row[2] || '',
          donde: row[3] || '',
          quien: row[4] || '',
          como: row[5] || '',
          cuando: row[6] ? new Date(row[6]).toISOString() : null,
          cuanto: row[7] || '',
          estado: row[8] || 'pendiente',
          fechaCreacion: row[9] ? new Date(row[9]).toISOString() : new Date().toISOString(),
          autor: row[10] || 'Sistema'
        };

        acciones.push(accion);
      }
    }

    console.log(`Acciones encontradas para Ciclo:${cicloId}, Causa:${idCausa}:`, acciones.length);
    return acciones;

  } catch (error) {
    console.error('Error en getAccionesCausa:', error);
    return [];
  }
}

function guardarAccionCausa(accionData) {
  try {
    console.log('Guardando acción 5W+2H:', accionData);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('AccionesCausa');

    // Si no existe la hoja, crearla con nuevas columnas
    if (!sheet) {
      console.log('Creando nueva hoja AccionesCausa con estructura 5W+2H');
      sheet = ss.insertSheet('AccionesCausa');

      const headers = [
        'Ciclo ID',
        'ID Causa',
        'QUÉ (Descripción)',
        'DÓNDE',
        'QUIÉN (Responsable)',
        'CÓMO',
        'CUÁNDO (Fecha Límite)',
        'CUÁNTO (Recursos/Costo)',
        'Estado',
        'Fecha Creación',
        'Autor'
      ];

      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Crear nueva fila con todos los campos
    const newRow = [
      accionData.cicloId,
      parseInt(accionData.idCausa),
      accionData.que || '',
      accionData.donde || '',
      accionData.quien || '',
      accionData.como || '',
      accionData.cuando || '',
      accionData.cuanto || '',
      accionData.estado || 'pendiente',
      new Date(),
      accionData.autor || 'Usuario'
    ];

    console.log('Nueva fila 5W+2H:', newRow);

    // Agregar fila
    sheet.appendRow(newRow);

    return {
      success: true,
      message: 'Acción 5W+2H guardada correctamente',
      accion: accionData
    };

  } catch (error) {
    console.error('Error en guardarAccionCausa:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}


function eliminarAccionCausa(cicloId, idCausa, indexAccion) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccionesCausa');
    if (!sheet) return { success: false, message: 'Hoja no encontrada' };

    const data = sheet.getDataRange().getValues();
    let contador = 0;
    let filaAEliminar = -1;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === cicloId && row[1] == idCausa) {
        if (contador == indexAccion) {
          filaAEliminar = i + 1; // +1 porque getValues() empieza en 0 pero las filas en 1
          break;
        }
        contador++;
      }
    }

    if (filaAEliminar > 0) {
      sheet.deleteRow(filaAEliminar);
      return { success: true, message: 'Acción eliminada' };
    }

    return { success: false, message: 'Acción no encontrada' };
  } catch (error) {
    console.error('Error en eliminarAccionCausa:', error);
    return { success: false, message: error.toString() };
  }
}

function getContadorAccionesCausa(cicloId, idCausa) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccionesCausa');
    if (!sheet) return 0;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;

    let contador = 0;
    const idCausaNum = parseInt(idCausa);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = row[0] || '';
      const rowIdCausa = parseInt(row[1]) || 0;

      if (rowCicloId.toString() === cicloId.toString() && rowIdCausa === idCausaNum) {
        contador++;
      }
    }

    return contador;
  } catch (error) {
    console.error('Error en getContadorAccionesCausa:', error);
    return 0;
  }
}

function crearHojaAccionesCausa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('AccionesCausa');

  const headers = [
    'Ciclo ID',
    'ID Causa',
    'Descripción',
    'Responsable',
    'Fecha Límite',
    'Estado',
    'Fecha Creación',
    'Autor'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  return sheet;
}

// ========== FUNCIONES SEGUIMIENTO POR ACCIÓN-CAUSA ==========

/**
 * Obtiene el historial de seguimiento de una acción específica dentro de una causa
 */
function getHistorialAccionCausa(cicloId, causaIndex) {
  try {
    console.log('🔍 Buscando historial para acción-causa - Ciclo:', cicloId, 'CausaIndex:', causaIndex);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('CICLOS_HISTORIAL');

    if (!sheet) {
      console.log('📭 Hoja de historial no encontrada');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    console.log('📊 Datos en hoja de historial:', data.length, 'filas');

    if (data.length <= 1) {
      console.log('📭 Hoja de historial vacía o solo encabezados');
      return [];
    }

    const historial = [];
    const causaIndexNum = parseInt(causaIndex);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();
      const rowCausaIndex = row[1] !== undefined ? parseInt(row[1]) : -1;

      // Buscar registros para este ciclo Y esta causa específica
      // Si row[1] está vacío, significa que es un seguimiento general del ciclo
      // Si row[1] tiene valor, es un seguimiento específico de una causa
      if (rowCicloId === cicloId && rowCausaIndex === causaIndexNum) {
        const registro = {
          cicloId: rowCicloId,
          causaIndex: rowCausaIndex,
          fecha: row[2] ? row[2].toISOString() : new Date().toISOString(),
          estado: String(row[3] || ''),
          comentario: String(row[4] || ''),
          autor: String(row[5] || 'Sistema')
        };

        console.log('✅ Registro encontrado para acción-causa:', registro);
        historial.push(registro);
      }
    }

    console.log('📋 Total registros encontrados para acción-causa:', historial.length);

    // Ordenar por fecha descendente (más reciente primero)
    historial.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

    return historial;

  } catch (error) {
    console.error('💥 Error crítico en getHistorialAccionCausa:', error);
    return [];
  }
}

/**
 * Guarda un seguimiento para una acción específica dentro de una causa
 */
function guardarSeguimientoAccionCausa(seguimiento) {
  try {
    console.log('💾 Guardando seguimiento para acción-causa:', seguimiento);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Verificar y preparar la hoja CICLOS_HISTORIAL
    let historialSheet = ss.getSheetByName('CICLOS_HISTORIAL');
    if (!historialSheet) {
      console.log('📝 Creando nueva hoja de historial');
      historialSheet = ss.insertSheet('CICLOS_HISTORIAL');
      // Actualizar encabezados para incluir Causa Index
      historialSheet.appendRow(['Ciclo ID', 'Causa Index', 'Fecha', 'Estado', 'Comentario', 'Autor', 'Tipo']);
      historialSheet.getRange(1, 1, 1, 7).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
    } else {
      // Verificar si la hoja tiene los encabezados correctos
      const headers = historialSheet.getRange(1, 1, 1, historialSheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Causa Index')) {
        // Agregar columna Causa Index si no existe
        historialSheet.insertColumnAfter(1);
        historialSheet.getRange(1, 2).setValue('Causa Index');
        historialSheet.getRange(1, 1, 1, historialSheet.getLastColumn()).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
      }
    }

    const fechaActual = new Date();

    // Agregar registro al historial
    historialSheet.appendRow([
      seguimiento.cicloId,
      parseInt(seguimiento.causaIndex),
      fechaActual,
      seguimiento.estado,
      seguimiento.comentario,
      seguimiento.autor || 'Usuario desconocido',
      'Acción-Causa' // Tipo para diferenciar de seguimientos generales del ciclo
    ]);

    // Formatear la fecha en la hoja
    const lastRow = historialSheet.getLastRow();
    historialSheet.getRange(lastRow, 3).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('✅ Seguimiento de acción-causa guardado en CICLOS_HISTORIAL, fila:', lastRow);

    // 2. También necesitamos modificar la función getHistorialCiclo para que ignore registros de acción-causa
    // (Eso se hará en esa función)

    return {
      success: true,
      message: 'Seguimiento guardado correctamente',
      detalles: {
        historialRow: lastRow,
        fecha: fechaActual.toISOString()
      }
    };

  } catch (error) {
    console.error('❌ Error en guardarSeguimientoAccionCausa:', error);
    return {
      success: false,
      message: 'Error al guardar seguimiento: ' + error.toString()
    };
  }
}

/**
 * Obtiene el estado actual de una acción específica
 */
function getEstadoAccionCausa(cicloId, causaIndex) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('CICLOS_HISTORIAL');

    if (!sheet) return 'Abierto';

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 'Abierto';

    const causaIndexNum = parseInt(causaIndex);
    let ultimoEstado = 'Abierto';

    // Buscar el último estado registrado para esta acción
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();
      const rowCausaIndex = row[1] !== undefined ? parseInt(row[1]) : -1;

      // Solo considerar registros específicos de esta causa
      if (rowCicloId === cicloId && rowCausaIndex === causaIndexNum) {
        ultimoEstado = String(row[3] || 'Abierto');
        break;
      }
    }

    return ultimoEstado;

  } catch (error) {
    console.error('Error en getEstadoAccionCausa:', error);
    return 'Abierto';
  }
}

/**
 * Guarda un seguimiento de causa en la hoja "Seguimiento_Causa"
 */
function guardarSeguimientoCausa(seguimientoData) {
  try {
    console.log('💾 Guardando seguimiento de causa:', seguimientoData);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Seguimiento_Causa');
    
    // Crear hoja si no existe
    if (!sheet) {
      console.log('📝 Creando nueva hoja: Seguimiento_Causa');
      sheet = ss.insertSheet('Seguimiento_Causa');
      
      // Configurar encabezados
      const headers = [
        'Ciclo ID',
        'Causa Index',
        'Causa Texto',
        'Fecha Registro',
        'Estado',
        'Comentario',
        'Fecha Próxima Revisión',
        'Autor',
        'Tipo'
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Formatear encabezados
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#0f307f');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // Congelar primera fila
      sheet.setFrozenRows(1);
      
      // Ajustar ancho de columnas
      sheet.autoResizeColumns(1, headers.length);
    }
    
    const fechaActual = new Date();
    
    // Preparar fecha próxima revisión si existe
    let fechaProxima = '';
    if (seguimientoData.fechaProximaRevision) {
      try {
        fechaProxima = new Date(seguimientoData.fechaProximaRevision);
      } catch (e) {
        console.error('Error parseando fecha próxima:', e);
        fechaProxima = '';
      }
    }
    
    // Preparar datos para la fila
    const rowData = [
      seguimientoData.cicloId,
      seguimientoData.causaIndex,
      seguimientoData.causaTexto || '',
      fechaActual, // Fecha como objeto Date
      seguimientoData.estado || 'Pendiente',
      seguimientoData.comentario || '',
      fechaProxima, // Fecha próxima (puede ser vacía)
      seguimientoData.autor || 'Usuario',
      seguimientoData.tipo || 'causa'
    ];
    
    console.log('📝 Datos a guardar:', rowData);
    
    // Agregar nueva fila
    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Formatear fechas en la hoja
    sheet.getRange(lastRow, 4).setNumberFormat('dd/mm/yyyy HH:mm'); // Columna fecha registro
    if (fechaProxima) {
      sheet.getRange(lastRow, 7).setNumberFormat('yyyy-mm-dd'); // Columna fecha próxima
    }
    
    // Formatear celda de estado
    const estadoCell = sheet.getRange(lastRow, 5);
    const estado = seguimientoData.estado || 'Pendiente';
    
    const estadoColors = {
        'Abierto': '#6b7280',
        'En Progreso': '#3b82f6',
        'Implementado': '#f59e0b',
        'Cerrado': '#6b7280'
    };
    
    estadoCell.setBackground(estadoColors[estado] || '#f3f4f6');
    estadoCell.setFontWeight('bold');
    
    console.log('✅ Seguimiento guardado en fila:', lastRow);
    
    return {
      success: true,
      message: 'Seguimiento de causa guardado correctamente',
      rowNumber: lastRow
    };
    
  } catch (error) {
    console.error('❌ Error en guardarSeguimientoCausa:', error);
    return {
      success: false,
      message: 'Error al guardar seguimiento: ' + error.toString()
    };
  }
}


function getHistorialSeguimientoCausa(cicloId, causaIndex) {
  try {
    console.log('🔍 BUSCANDO HISTORIAL - Ciclo:', cicloId, 'Causa:', causaIndex);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Seguimiento_Causa');
    
    if (!sheet) {
      console.log('📭 Hoja Seguimiento_Causa no encontrada');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    console.log('📊 Total filas en hoja:', lastRow);
    
    if (lastRow <= 1) {
      console.log('📭 Solo encabezados, sin datos');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    console.log('📋 Datos obtenidos, filas:', data.length);
    
    // Mostrar las primeras filas para diagnóstico
    console.log('Primeras 3 filas de datos:');
    for (let i = 0; i < Math.min(3, data.length); i++) {
      console.log(`Fila ${i}:`, data[i]);
    }
    
    const historial = [];
    const causaIndexNum = parseInt(causaIndex);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Validar fila no vacía
      if (!row[0] && !row[1]) continue;
      
      const rowCicloId = String(row[0] || '').trim();
      const rowCausaIndex = parseInt(row[1]) || -1;
      
      console.log(`\nFila ${i + 1}:`);
      console.log('- rowCicloId:', rowCicloId, 'buscando:', cicloId);
      console.log('- rowCausaIndex:', rowCausaIndex, 'buscando:', causaIndexNum);
      console.log('- Coincide ciclo?', rowCicloId === cicloId);
      console.log('- Coincide causa?', rowCausaIndex === causaIndexNum);
      console.log('- Fecha cruda (row[3]):', row[3], 'tipo:', typeof row[3]);
      
      if (rowCicloId === cicloId && rowCausaIndex === causaIndexNum) {
        console.log('✅ ✅ ✅ REGISTRO COINCIDE!');
        
        // CONVERTIR FECHA DE TEXTO A DATE
        let fecha = null;
        try {
          if (row[3] instanceof Date) {
            fecha = row[3];
          } else if (typeof row[3] === 'string' && row[3].trim()) {
            // Intentar parsear fecha en formato dd/MM/yyyy HH:mm
            const dateParts = row[3].split(' ');
            const dateStr = dateParts[0]; // "13/01/2026"
            const timeStr = dateParts[1] || '00:00'; // "11:30"
            
            const [day, month, year] = dateStr.split('/');
            const [hour, minute] = timeStr.split(':');
            
            fecha = new Date(year, month - 1, day, hour, minute);
            console.log('📅 Fecha convertida:', fecha);
          }
        } catch (e) {
          console.error('Error convirtiendo fecha:', e);
          fecha = new Date(); // Fecha actual como fallback
        }
        
        if (!fecha || isNaN(fecha.getTime())) {
          fecha = new Date(); // Fecha actual si no se pudo convertir
        }
        
        const registro = {
          cicloId: rowCicloId,
          causaIndex: rowCausaIndex,
          causaTexto: String(row[2] || ''),
          fecha: fecha.toISOString(), // Usar fecha convertida
          estado: String(row[4] || 'Pendiente'),
          comentario: String(row[5] || ''),
          fechaProximaRevision: row[6] ? 
            (typeof row[6] === 'string' ? row[6] : Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'dd/MM/yyyy')) : '',
          autor: String(row[7] || 'Usuario'),
          tipo: String(row[8] || 'causa')
        };
        
        console.log('📝 Registro creado:', registro);
        historial.push(registro);
      }
    }
    
    console.log('🎯 Total registros encontrados:', historial.length);
    
    // Ordenar por fecha descendente (más reciente primero)
    historial.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    
    return historial;
    
  } catch (error) {
    console.error('💥 ERROR en getHistorialSeguimientoCausa:', error);
    console.error('Stack trace:', error.stack);
    return [];
  }
}

/**
 * Obtiene el estado actual de una causa
 */
function getEstadoActualCausa(cicloId, causaIndex) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Seguimiento_Causa');

    if (!sheet) return 'Pendiente';

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 'Pendiente';

    const causaIndexNum = parseInt(causaIndex);
    let ultimoEstado = 'Pendiente';
    let ultimaFecha = new Date(0);

    // Buscar el último estado registrado
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();
      const rowCausaIndex = parseInt(row[1]) || -1;

      if (rowCicloId === cicloId && rowCausaIndex === causaIndexNum) {
        const fechaRegistro = row[3] instanceof Date ? row[3] : new Date(0);

        if (fechaRegistro > ultimaFecha) {
          ultimaFecha = fechaRegistro;
          ultimoEstado = String(row[4] || 'Pendiente');
        }
      }
    }

    return ultimoEstado;

  } catch (error) {
    console.error('Error en getEstadoActualCausa:', error);
    return 'Pendiente';
  }
}