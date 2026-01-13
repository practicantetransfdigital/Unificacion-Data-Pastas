//BACKEND DE INDICADORES O ESTADISTICAS PARA N2

function getN2StatsAnioMesProceso() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_N2');

    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const fechaIndex = headers.indexOf('Fecha de Registro');
    const procesoIndex = headers.indexOf('Proceso');
    const estadoIndex = headers.indexOf('Estado');

    if (fechaIndex === -1 || procesoIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }

    const stats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fechaStr = row[fechaIndex];
      const proceso = row[procesoIndex] || 'Sin proceso';
      const estado = row[estadoIndex] || 'Sin estado';

      let fecha;
      try {
        fecha = new Date(fechaStr);
      } catch (e) {
        continue;
      }

      if (!fecha || isNaN(fecha.getTime())) {
        continue;
      }

      const year = fecha.getFullYear();
      const month = fecha.getMonth() + 1;
      const mesNombre = getMonthName(month);
      const key = `${year}-${month.toString().padStart(2, '0')}-${proceso}`;

      if (!stats[key]) {
        stats[key] = {
          year: year,
          month: month,
          mesNombre: mesNombre,
          proceso: proceso,
          completado: 0,
          enProgreso: 0,
          enRevision: 0,
          pendiente: 0,
          total: 0
        };
      }

      if (estado === 'Completado') {
        stats[key].completado++;
      } else if (estado === 'En Progreso') {
        stats[key].enProgreso++;
      } else if (estado === 'En Revision') {
        stats[key].enRevision++;
      } else if (estado === 'Pendiente') {
        stats[key].pendiente++;
      }

      stats[key].total++;
    }

    const tableData = [];
    const statsByYear = {};
    const statsByYearMonth = {};
    let totalGlobal = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };

    Object.keys(stats).sort().forEach(key => {
      const stat = stats[key];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        ano: stat.year,
        mes: stat.mesNombre,
        proceso: stat.proceso,
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%'
      });

      const yearMonthKey = `${stat.year}-${stat.month}`;
      if (!statsByYearMonth[yearMonthKey]) {
        statsByYearMonth[yearMonthKey] = { year: stat.year, mes: stat.mesNombre, completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };
      }
      statsByYearMonth[yearMonthKey].completado += stat.completado;
      statsByYearMonth[yearMonthKey].enProgreso += stat.enProgreso;
      statsByYearMonth[yearMonthKey].enRevision += stat.enRevision;
      statsByYearMonth[yearMonthKey].pendiente += stat.pendiente;
      statsByYearMonth[yearMonthKey].total += stat.total;

      if (!statsByYear[stat.year]) {
        statsByYear[stat.year] = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };
      }
      statsByYear[stat.year].completado += stat.completado;
      statsByYear[stat.year].enProgreso += stat.enProgreso;
      statsByYear[stat.year].enRevision += stat.enRevision;
      statsByYear[stat.year].pendiente += stat.pendiente;
      statsByYear[stat.year].total += stat.total;

      totalGlobal.completado += stat.completado;
      totalGlobal.enProgreso += stat.enProgreso;
      totalGlobal.enRevision += stat.enRevision;
      totalGlobal.pendiente += stat.pendiente;
      totalGlobal.total += stat.total;
    });

    Object.keys(statsByYearMonth).sort().forEach(key => {
      const stat = statsByYearMonth[key];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        ano: `Total ${stat.mes}`,
        mes: '',
        proceso: '',
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%',
        isMonthTotal: true
      });
    });

    Object.keys(statsByYear).sort().forEach(year => {
      const stat = statsByYear[year];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        ano: `Total ${year}`,
        mes: '',
        proceso: '',
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%',
        isYearTotal: true
      });
    });

    const globalPorcentaje = totalGlobal.total > 0 ?
      ((totalGlobal.completado / totalGlobal.total) * 100).toFixed(2) : '0.00';

    tableData.push({
      ano: '',
      mes: '',
      proceso: 'Suma total',
      completado: totalGlobal.completado,
      enProgreso: totalGlobal.enProgreso,
      enRevision: totalGlobal.enRevision,
      pendiente: totalGlobal.pendiente,
      total: totalGlobal.total,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });

    return {
      success: true,
      data: tableData,
      summary: {
        total: totalGlobal.total,
        completado: totalGlobal.completado,
        enProgreso: totalGlobal.enProgreso,
        enRevision: totalGlobal.enRevision,
        pendiente: totalGlobal.pendiente,
        porcentajeGestion: globalPorcentaje
      }
    };

  } catch (error) {
    console.error('Error en getN2StatsAnioMesProceso:', error);
    return { success: false, message: error.toString() };
  }
}

function getN2StatsProceso() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_N2');

    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const procesoIndex = headers.indexOf('Proceso');
    const estadoIndex = headers.indexOf('Estado');

    if (procesoIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }

    const stats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const proceso = row[procesoIndex] || 'Sin proceso';
      const estado = row[estadoIndex] || 'Sin estado';

      if (!stats[proceso]) {
        stats[proceso] = {
          completado: 0,
          enProgreso: 0,
          enRevision: 0,
          pendiente: 0,
          total: 0
        };
      }

      if (estado === 'Completado') {
        stats[proceso].completado++;
      } else if (estado === 'En Progreso') {
        stats[proceso].enProgreso++;
      } else if (estado === 'En Revision') {
        stats[proceso].enRevision++;
      } else if (estado === 'Pendiente') {
        stats[proceso].pendiente++;
      }

      stats[proceso].total++;
    }

    const tableData = [];
    let totalGlobal = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };

    Object.keys(stats).sort().forEach(proceso => {
      const stat = stats[proceso];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        proceso: proceso,
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%'
      });

      totalGlobal.completado += stat.completado;
      totalGlobal.enProgreso += stat.enProgreso;
      totalGlobal.enRevision += stat.enRevision;
      totalGlobal.pendiente += stat.pendiente;
      totalGlobal.total += stat.total;
    });

    const globalPorcentaje = totalGlobal.total > 0 ?
      ((totalGlobal.completado / totalGlobal.total) * 100).toFixed(2) : '0.00';

    tableData.push({
      proceso: 'Suma total',
      completado: totalGlobal.completado,
      enProgreso: totalGlobal.enProgreso,
      enRevision: totalGlobal.enRevision,
      pendiente: totalGlobal.pendiente,
      total: totalGlobal.total,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });

    return {
      success: true,
      data: tableData
    };

  } catch (error) {
    console.error('Error en getN2StatsProceso:', error);
    return { success: false, message: error.toString() };
  }
}

function getN2StatsZonaProceso() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_N2');

    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const procesoIndex = headers.indexOf('Proceso');
    const zonaProcesoIndex = headers.indexOf('ZonaProceso');
    const estadoIndex = headers.indexOf('Estado');

    if (procesoIndex === -1 || zonaProcesoIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }

    const stats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const proceso = row[procesoIndex] || 'Sin proceso';
      const zonaProceso = row[zonaProcesoIndex] || 'Sin zona';
      const estado = row[estadoIndex] || 'Sin estado';
      const key = `${proceso}-${zonaProceso}`;

      if (!stats[key]) {
        stats[key] = {
          proceso: proceso,
          zonaProceso: zonaProceso,
          completado: 0,
          enProgreso: 0,
          enRevision: 0,
          pendiente: 0,
          total: 0
        };
      }

      if (estado === 'Completado') {
        stats[key].completado++;
      } else if (estado === 'En Progreso') {
        stats[key].enProgreso++;
      } else if (estado === 'En Revision') {
        stats[key].enRevision++;
      } else if (estado === 'Pendiente') {
        stats[key].pendiente++;
      }

      stats[key].total++;
    }

    const tableData = [];
    const statsByProceso = {};
    let totalGlobal = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };

    Object.keys(stats).sort().forEach(key => {
      const stat = stats[key];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        proceso: stat.proceso,
        zonaProceso: stat.zonaProceso,
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%'
      });

      if (!statsByProceso[stat.proceso]) {
        statsByProceso[stat.proceso] = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };
      }
      statsByProceso[stat.proceso].completado += stat.completado;
      statsByProceso[stat.proceso].enProgreso += stat.enProgreso;
      statsByProceso[stat.proceso].enRevision += stat.enRevision;
      statsByProceso[stat.proceso].pendiente += stat.pendiente;
      statsByProceso[stat.proceso].total += stat.total;

      totalGlobal.completado += stat.completado;
      totalGlobal.enProgreso += stat.enProgreso;
      totalGlobal.enRevision += stat.enRevision;
      totalGlobal.pendiente += stat.pendiente;
      totalGlobal.total += stat.total;
    });

    Object.keys(statsByProceso).sort().forEach(proceso => {
      const stat = statsByProceso[proceso];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        proceso: `Total ${proceso}`,
        zonaProceso: '',
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%',
        isProcesoTotal: true
      });
    });

    const globalPorcentaje = totalGlobal.total > 0 ?
      ((totalGlobal.completado / totalGlobal.total) * 100).toFixed(2) : '0.00';

    tableData.push({
      proceso: '',
      zonaProceso: 'Suma total',
      completado: totalGlobal.completado,
      enProgreso: totalGlobal.enProgreso,
      enRevision: totalGlobal.enRevision,
      pendiente: totalGlobal.pendiente,
      total: totalGlobal.total,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });

    return {
      success: true,
      data: tableData
    };

  } catch (error) {
    console.error('Error en getN2StatsZonaProceso:', error);
    return { success: false, message: error.toString() };
  }
}

function getN2StatsProcesoResponsable() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_N2');

    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const procesoResponsableIndex = headers.indexOf('Proceso Responsable');
    const estadoIndex = headers.indexOf('Estado');

    if (procesoResponsableIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }

    const stats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const procesoResponsable = row[procesoResponsableIndex] || 'Sin proceso responsable';
      const estado = row[estadoIndex] || 'Sin estado';

      if (!stats[procesoResponsable]) {
        stats[procesoResponsable] = {
          completado: 0,
          enProgreso: 0,
          enRevision: 0,
          pendiente: 0,
          total: 0
        };
      }

      if (estado === 'Completado') {
        stats[procesoResponsable].completado++;
      } else if (estado === 'En Progreso') {
        stats[procesoResponsable].enProgreso++;
      } else if (estado === 'En Revision') {
        stats[procesoResponsable].enRevision++;
      } else if (estado === 'Pendiente') {
        stats[procesoResponsable].pendiente++;
      }

      stats[procesoResponsable].total++;
    }

    const tableData = [];
    let totalGlobal = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };

    Object.keys(stats).sort().forEach(procesoResponsable => {
      const stat = stats[procesoResponsable];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        procesoResponsable: procesoResponsable,
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%'
      });

      totalGlobal.completado += stat.completado;
      totalGlobal.enProgreso += stat.enProgreso;
      totalGlobal.enRevision += stat.enRevision;
      totalGlobal.pendiente += stat.pendiente;
      totalGlobal.total += stat.total;
    });

    const globalPorcentaje = totalGlobal.total > 0 ?
      ((totalGlobal.completado / totalGlobal.total) * 100).toFixed(2) : '0.00';

    tableData.push({
      procesoResponsable: 'Suma total',
      completado: totalGlobal.completado,
      enProgreso: totalGlobal.enProgreso,
      enRevision: totalGlobal.enRevision,
      pendiente: totalGlobal.pendiente,
      total: totalGlobal.total,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });

    return {
      success: true,
      data: tableData
    };

  } catch (error) {
    console.error('Error en getN2StatsProcesoResponsable:', error);
    return { success: false, message: error.toString() };
  }
}

function getN2StatsLiderResponsable() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_N2');

    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const liderIndex = headers.indexOf('Líder Responsable');
    const estadoIndex = headers.indexOf('Estado');

    if (liderIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }

    const stats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const lider = row[liderIndex] || 'Sin líder';
      const estado = row[estadoIndex] || 'Sin estado';

      if (!stats[lider]) {
        stats[lider] = {
          completado: 0,
          enProgreso: 0,
          enRevision: 0,
          pendiente: 0,
          total: 0
        };
      }

      if (estado === 'Completado') {
        stats[lider].completado++;
      } else if (estado === 'En Progreso') {
        stats[lider].enProgreso++;
      } else if (estado === 'En Revision') {
        stats[lider].enRevision++;
      } else if (estado === 'Pendiente') {
        stats[lider].pendiente++;
      }

      stats[lider].total++;
    }

    const tableData = [];
    let totalGlobal = { completado: 0, enProgreso: 0, enRevision: 0, pendiente: 0, total: 0 };

    Object.keys(stats).sort().forEach(lider => {
      const stat = stats[lider];
      const porcentaje = stat.total > 0 ?
        ((stat.completado / stat.total) * 100).toFixed(2) : '0.00';

      tableData.push({
        liderResponsable: lider,
        completado: stat.completado,
        enProgreso: stat.enProgreso,
        enRevision: stat.enRevision,
        pendiente: stat.pendiente,
        total: stat.total,
        porcentaje: porcentaje + '%'
      });

      totalGlobal.completado += stat.completado;
      totalGlobal.enProgreso += stat.enProgreso;
      totalGlobal.enRevision += stat.enRevision;
      totalGlobal.pendiente += stat.pendiente;
      totalGlobal.total += stat.total;
    });

    const globalPorcentaje = totalGlobal.total > 0 ?
      ((totalGlobal.completado / totalGlobal.total) * 100).toFixed(2) : '0.00';

    tableData.push({
      liderResponsable: 'Suma total',
      completado: totalGlobal.completado,
      enProgreso: totalGlobal.enProgreso,
      enRevision: totalGlobal.enRevision,
      pendiente: totalGlobal.pendiente,
      total: totalGlobal.total,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });

    return {
      success: true,
      data: tableData
    };

  } catch (error) {
    console.error('Error en getN2StatsLiderResponsable:', error);
    return { success: false, message: error.toString() };
  }
}