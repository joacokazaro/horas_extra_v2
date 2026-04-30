const CFG = {
  SHEET_CONFIG: 'DATOS SCRIPT',
  SHEET_VALORES: 'VALORES HORAS',
  TOPE_BASE_MENSUAL: 192,
  TZ: Session.getScriptTimeZone() || 'America/Argentina/Cordoba',
  AUTHORIZED_EXECUTORS: [
    'francisco.savid@kazaro.com.ar',
    'lautaro.suarez@kazaro.com.ar',
    'joaquin.rojas@kazaro.com.ar',
    'servicio@kazaro.com.ar',
  ],

  SERVICE_CATALOG_URL: 'https://docs.google.com/spreadsheets/d/1Mo_E4ZOngzvu-yBWl0fM05Xz5eeYq1C0CEWP0oL17Hc/edit',
  SERVICE_CATALOG_SHEET: 'Lista de Servicios',
  SERVICE_CATALOG_SERVICE_COL: 4,   // D
  SERVICE_CATALOG_TYPE_COL: 29,     // AC

  RESULT_HEADERS: [
    'LEG', 'NOMBRE', 'APELLIDO', 'TOPE', 'HS SOLICITADAS', 'HS TEORICAS',
    'HS REALES', 'CONTROL OP', 'DIFERENCIA', 'HS 100', 'HS 50', 'HS NORMALES', 'SUPERVISOR'
  ],

  CONFIG_CELLS: {
    urlNomina: 'B1',
    hojaNomina: 'B2',
    mes: 'B3',
    anio: 'B4',
    desdeFila: 'B5',
    hastaFila: 'B6',
    ultimaEjecucion: 'B7',
    estado: 'B8',
    mensaje: 'B9',
  },
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Horas Extra')
    .addItem('Validar configuración', 'validarConfiguracion')
    .addSeparator()
    .addItem('Generar resumen final', 'generarResultadosMes')
    .addToUi();
}

function validarConfiguracion() {
  try {
    const ctx = getContexto_();
    const resumenFeriados = resumirConfiguracionFeriados_(ctx.feriadosConfig);
    const mensaje = `Configuración válida para ${ctx.mes}/${ctx.anio} | ${resumenFeriados}`;
    setEstado_('OK', mensaje);
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (error) {
    setEstado_('ERROR', error.message || String(error));
    throw error;
  }
}

function generarResultadosMes() {
  if (!usuarioAutorizadoParaEjecutar_()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No tenes permisos para ejecutar este boton');
    return;
  }

  try {
    const ctx = getContexto_();
    const datos = cargarDatosBase_(ctx);
    const novedadesPorLegajo = agruparNovedadesPorLegajo_(datos.novedades, ctx);
    const nominaPorLegajo = indexarNominaPorLegajo_(datos.nomina);
    const horasRealesPorLegajo = resumirHorasRealesPorLegajo_(datos.horasReales);
    const tiposServicio = cargarMapaTiposServicio_();

    const resultado = construirResumenPorLegajo_(
      ctx,
      datos.operativa,
      novedadesPorLegajo,
      nominaPorLegajo,
      horasRealesPorLegajo,
      tiposServicio,
      datos.valoresHora
    );
    const resumen = resultado.filas;
    const mensaje = construirMensajeEjecucion_(resultado, datos);

    escribirDesgloseEnOperativa_(ctx.hojaOperativa, resultado.operativa, novedadesPorLegajo);
    escribirResultados_(ctx.hojaResultados, resumen);
    setEstado_('OK', mensaje);
    SpreadsheetApp.getActiveSpreadsheet().toast(`${resumen.length} legajos procesados`, 'Horas Extra', 5);
    SpreadsheetApp.getUi().alert(`Proceso finalizado: ${mensaje}`);
  } catch (error) {
    setEstado_('ERROR', error.message || String(error));
    SpreadsheetApp.getUi().alert(`Error al generar resultados: ${error.message || String(error)}`);
    throw error;
  }
}

function getContexto_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfig = ss.getSheetByName(CFG.SHEET_CONFIG);
  if (!hojaConfig) throw new Error(`No existe la hoja ${CFG.SHEET_CONFIG}`);

  const config = leerConfiguracion_(hojaConfig);
  const feriadosConfig = leerConfiguracionFeriados_(hojaConfig);
  const urlNomina = String(config.urlNomina || '').trim();
  const hojaNominaNombre = String(config.hojaNomina || '').trim();
  const mes = Number(config.mes);
  const anio = Number(config.anio);
  const desdeFila = Number(config.desdeFila || 2);
  const hastaFilaRaw = config.hastaFila;
  const hastaFila = String(hastaFilaRaw || '').trim() ? Number(hastaFilaRaw) : null;

  if (!urlNomina) throw new Error('Falta URL NOMINA en DATOS SCRIPT');
  if (!hojaNominaNombre) throw new Error('Falta HOJA NOMINA en DATOS SCRIPT');
  if (!mes || mes < 1 || mes > 12) throw new Error('MES inválido en DATOS SCRIPT');
  if (!anio || anio < 2024) throw new Error('AÑO inválido en DATOS SCRIPT');
  if (!desdeFila || desdeFila < 2) throw new Error('DESDE FILA inválido en DATOS SCRIPT');
  if (hastaFila !== null && (!hastaFila || hastaFila < desdeFila)) {
    throw new Error('HASTA FILA inválido en DATOS SCRIPT');
  }

  const nombres = resolverNombresHojas_(mes, anio);
  const hojaOperativa = obtenerHojaPorAlias_(ss, nombres.horasExtra);
  const hojaNovedades = obtenerHojaPorAlias_(ss, nombres.novedades);
  const hojaResultados = obtenerHojaPorAlias_(ss, nombres.resultados);
  const hojaValores = ss.getSheetByName(CFG.SHEET_VALORES);

  if (!hojaOperativa) throw new Error(`No existe ninguna hoja operativa válida: ${nombres.horasExtra.join(' / ')}`);
  if (!hojaNovedades) throw new Error(`No existe ninguna hoja de novedades válida: ${nombres.novedades.join(' / ')}`);
  if (!hojaResultados) throw new Error(`No existe ninguna hoja de resultados válida: ${nombres.resultados.join(' / ')}`);
  if (!hojaValores) throw new Error(`No existe la hoja ${CFG.SHEET_VALORES}`);

  const archivoNomina = SpreadsheetApp.openByUrl(urlNomina);
  const hojaNomina = archivoNomina.getSheetByName(hojaNominaNombre);
  const hojaHorasReales = archivoNomina.getSheetByName(hojaNominaNombre);
  if (!hojaNomina) throw new Error(`No existe la hoja ${hojaNominaNombre} en la nómina`);
  if (!hojaHorasReales) throw new Error(`No existe la hoja ${hojaNominaNombre} en el archivo externo`);

  return {
    ss,
    hojaConfig,
    hojaOperativa,
    hojaNovedades,
    hojaResultados,
    hojaValores,
    hojaNomina,
    hojaHorasReales,
    mes,
    anio,
    desdeFila,
    hastaFila,
    feriadosConfig,
  };
}

function resolverNombresHojas_(mes, anio) {
  const mesTexto = nombreMes_(mes);
  const anioCorto = String(anio).slice(-2);

  return {
    novedades: [`Novedades ${mesTexto} ${anioCorto}`, 'Novedades'],
    horasExtra: [`Horas Extra ${mesTexto} ${anioCorto}`, 'Horas Extra'],
    resultados: [`Resultados ${mesTexto} ${anioCorto}`, 'Resultados'],
  };
}

function obtenerHojaPorAlias_(ss, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const hoja = ss.getSheetByName(aliases[i]);
    if (hoja) return hoja;
  }

  return null;
}

function cargarDatosBase_(ctx) {
  return {
    operativa: ctx.hojaOperativa.getDataRange().getValues(),
    novedades: ctx.hojaNovedades.getDataRange().getValues(),
    nomina: ctx.hojaNomina.getDataRange().getValues(),
    valoresHora: ctx.hojaValores.getDataRange().getValues(),
    horasReales: ctx.hojaHorasReales.getDataRange().getValues(),
  };
}

function leerConfiguracion_(hojaConfig) {
  const lastRow = hojaConfig.getLastRow();
  const configPorEtiqueta = {};

  if (lastRow > 0) {
    const data = hojaConfig.getRange(1, 1, lastRow, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      const etiqueta = normalizarTexto_(data[i][0]);
      if (!etiqueta) continue;
      configPorEtiqueta[etiqueta] = data[i][1];
    }
  }

  return {
    urlNomina: valorConfiguracion_(configPorEtiqueta, ['url nomina', 'archivo nomina'], hojaConfig.getRange(CFG.CONFIG_CELLS.urlNomina).getValue()),
    hojaNomina: valorConfiguracion_(configPorEtiqueta, ['hoja nomina'], hojaConfig.getRange(CFG.CONFIG_CELLS.hojaNomina).getValue()),
    mes: valorConfiguracion_(configPorEtiqueta, ['mes'], hojaConfig.getRange(CFG.CONFIG_CELLS.mes).getValue()),
    anio: valorConfiguracion_(configPorEtiqueta, ['ano', 'año'], hojaConfig.getRange(CFG.CONFIG_CELLS.anio).getValue()),
    desdeFila: valorConfiguracion_(configPorEtiqueta, ['desde fila'], hojaConfig.getRange(CFG.CONFIG_CELLS.desdeFila).getValue()),
    hastaFila: valorConfiguracion_(configPorEtiqueta, ['hasta fila'], hojaConfig.getRange(CFG.CONFIG_CELLS.hastaFila).getValue()),
  };
}

function valorConfiguracion_(configPorEtiqueta, aliases, fallback) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizarTexto_(aliases[i]);
    if (configPorEtiqueta[key] !== undefined && String(configPorEtiqueta[key] || '').trim() !== '') {
      return configPorEtiqueta[key];
    }
  }

  return fallback;
}

function leerConfiguracionFeriados_(hojaConfig) {
  const data = hojaConfig.getDataRange().getValues();
  const feriadosPorTipo = {};
  const feriadosPorServicio = {
    exacto: {},
    simplificado: {},
    entradas: [],
  };
  let seccion = '';

  for (let i = 0; i < data.length; i++) {
    const etiqueta = String(data[i][0] || '').trim();
    const valor = data[i][1];
    const etiquetaNormalizada = normalizarTexto_(etiqueta);

    if (etiquetaNormalizada === 'tipo serv. / servicio') {
      seccion = 'tipo';
      continue;
    }

    if (etiquetaNormalizada === 'servicio cerrado por x motivo') {
      seccion = 'servicio';
      continue;
    }

    if (etiquetaNormalizada === 'dias feriados (fecha)' || etiquetaNormalizada === 'dia (fecha)') {
      continue;
    }

    if (!etiquetaNormalizada || valor === '' || valor === null) {
      continue;
    }

    const fechas = parseFechasMultiples_(valor);
    if (!fechas.length) continue;

    if (seccion === 'tipo') {
      const tipo = normalizarTipoServicio_(etiqueta);
      if (!tipo) continue;
      if (!feriadosPorTipo[tipo]) feriadosPorTipo[tipo] = new Set();
      fechas.forEach((fecha) => feriadosPorTipo[tipo].add(ymd_(fecha)));
      continue;
    }

    if (seccion === 'servicio') {
      const servicio = normalizarTexto_(etiqueta);
      const servicioSimplificado = simplificarTextoComparacion_(etiqueta);
      if (!servicio) continue;

      if (!feriadosPorServicio.exacto[servicio]) feriadosPorServicio.exacto[servicio] = new Set();
      fechas.forEach((fecha) => feriadosPorServicio.exacto[servicio].add(ymd_(fecha)));

      if (servicioSimplificado) {
        if (!feriadosPorServicio.simplificado[servicioSimplificado]) {
          feriadosPorServicio.simplificado[servicioSimplificado] = new Set();
        }
        fechas.forEach((fecha) => feriadosPorServicio.simplificado[servicioSimplificado].add(ymd_(fecha)));
      }

      fechas.forEach((fecha) => {
        feriadosPorServicio.entradas.push({
          normalizado: servicio,
          simplificado: servicioSimplificado,
          fecha: ymd_(fecha),
        });
      });
    }
  }

  return {
    porTipo: feriadosPorTipo,
    porServicio: feriadosPorServicio,
  };
}

function construirResumenPorLegajo_(ctx, operativa, novedadesPorLegajo, nominaPorLegajo, horasRealesPorLegajo, tiposServicio, valoresHoraData) {
  const seleccion = seleccionarFilasOperativas_(operativa, ctx);
  const rows = compactarFilasOperativas_(seleccion.rows);
  const acumulado = {};
  const valoresPorCategoria = indexarValoresHora_(valoresHoraData || []);

  seleccion.rowsCompactadas = rows;

  rows.forEach((row) => {
    const legajo = normalizarLegajo_(row[0]);
    if (!legajo) return;

    if (!acumulado[legajo]) {
      const persona = nominaPorLegajo[legajo] || {};
      const supervisor = String(row[5] || '').trim();
      const jornada = Number(persona.jornada || row[6] || 0);
      const categoria = String(persona.categoria || row[7] || '').trim();
      const servicioOrigen = String(row[8] || '').trim();
      const servicioAlternativo = String(persona.servicio || row[9] || '').trim();
      const servicio = resolverTipoServicio_(servicioOrigen, servicioAlternativo, tiposServicio);
      const ausencias = novedadesPorLegajo[legajo] || [];
      const nombreServicio = servicioOrigen || servicioAlternativo;
      const hsTeoricasBase = calcularHorasTeoricasBaseMensuales_(servicio, jornada, ctx.mes, ctx.anio);
      const hsTeoricas = calcularHorasTeoricasMensuales_(
        servicio,
        nombreServicio,
        jornada,
        ctx.mes,
        ctx.anio,
        ausencias,
        ctx.feriadosConfig
      );
      const tope = calcularTopeMensual_(hsTeoricasBase, hsTeoricas);

      acumulado[legajo] = {
        legajo,
        nombre: persona.nombre || row[2] || '',
        apellido: persona.apellido || row[1] || '',
        categoria,
        jornada,
        servicio,
        servicioNombre: nombreServicio,
        supervisor,
        tope,
        hsSolicitadas: 0,
        hsTeoricas,
        hsReales: numero_(horasRealesPorLegajo[legajo]),
        valorHoras: 0,
        controlOp: 0,
        diferencia: 0,
        hsNormales: 0,
        hs50: 0,
        hs100: 0,
        registros: [],
      };
    }

    const item = acumulado[legajo];
    const hsSolicitadas = numero_(row[14]); // O = Cantidad Horas
    const servicioRegistro = String(row[8] || '').trim() || item.servicioNombre;
    const tipoServicioRegistro = resolverTipoServicio_(servicioRegistro, servicioRegistro, tiposServicio) || item.servicio;
    item.hsSolicitadas += hsSolicitadas;
    item.registros.push({
      row: row.slice(),
      rowIndex: seleccion.desdeIndex + item.registros.length,
      fecha: row[10], // K = FECHA
      diaSemana: row[11], // L = Día Semana
      servicio: tipoServicioRegistro,
      servicioNombre: servicioRegistro,
      esEvento: esPedidoEvento_(row[20]), // U = Evento?
      horas: hsSolicitadas,
    });
  });

  Object.keys(acumulado).forEach((legajo) => {
    const item = acumulado[legajo];

    item.hsSolicitadas = redondear_(item.hsSolicitadas);
    item.hsReales = redondear_(item.hsReales);
    item.controlOp = redondear_(Math.max(0, Math.min(item.hsSolicitadas, item.hsReales - item.hsTeoricas)));
    item.diferencia = redondear_(item.hsSolicitadas - item.controlOp);

    // El cupo normal no depende del neto pagable sino del margen que queda
    // antes de alcanzar el tope mensual de la persona.
    const cupoNormal = redondear_(Math.max(0, item.tope - item.hsTeoricas));
    const desglose = desglosarHorasLegajo_({
      servicio: item.servicio,
      servicioNombre: item.servicioNombre,
      registros: item.registros,
      totalDisponible: item.controlOp,
      cupoNormal,
      feriadosConfig: ctx.feriadosConfig,
    });
    const valoresCategoria = valoresPorCategoria[normalizarTexto_(item.categoria)] || {};

    item.hsNormales = desglose.normal;
    item.hs50 = desglose.extra50;
    item.hs100 = desglose.extra100;
    item.valorHoras = calcularValorDesglose_(desglose, valoresCategoria);
    item.registrosDesglosados = desglose.registros;
  });

  const filas = Object.values(acumulado)
    .sort((a, b) => Number(a.legajo) - Number(b.legajo))
    .map((item) => ([
      item.legajo,
      item.nombre,
      item.apellido,
      item.tope,
      item.hsSolicitadas,
      item.hsTeoricas,
      item.hsReales,
      item.controlOp,
      item.diferencia,
      item.hs100,
      item.hs50,
      item.hsNormales,
      item.supervisor,
    ]));

  const operativaDesglosada = construirOperativaDesglosada_(operativa, seleccion, acumulado);

  return {
    filas,
    operativa: operativaDesglosada,
    diagnostico: {
      modoRango: seleccion.modo,
      filasConsideradas: rows.length,
      legajosProcesados: filas.length,
      desdeFilaUsada: seleccion.desdeFila,
      hastaFilaUsada: seleccion.hastaFila,
    },
  };
}

function seleccionarFilasOperativas_(operativa, ctx) {
  const desdeIndexConfigurado = Math.max(ctx.desdeFila - 1, 1);
  const hastaIndexConfigurado = ctx.hastaFila ? ctx.hastaFila : operativa.length;
  const rowsConfiguradas = operativa.slice(desdeIndexConfigurado, hastaIndexConfigurado);

  if (contarLegajosEnFilas_(rowsConfiguradas) > 0) {
    return {
      rows: rowsConfiguradas,
      modo: 'configurado',
      desdeFila: ctx.desdeFila,
      hastaFila: ctx.hastaFila || operativa.length,
      desdeIndex: desdeIndexConfigurado,
      hastaIndex: hastaIndexConfigurado,
    };
  }

  const rowsCompletas = operativa.slice(1);
  return {
    rows: rowsCompletas,
    modo: 'completo',
    desdeFila: 2,
    hastaFila: operativa.length,
    desdeIndex: 1,
    hastaIndex: operativa.length,
  };
}

function construirOperativaDesglosada_(operativa, seleccion, acumulado) {
  const salida = [operativa[0].slice()];
  const rowsAntes = operativa.slice(1, seleccion.desdeIndex);
  const rowsDespues = operativa.slice(seleccion.hastaIndex);
  const rowsDesglosar = seleccion.rowsCompactadas || seleccion.rows;

  for (let i = 0; i < rowsAntes.length; i++) {
    salida.push(limpiarColumnasDesglose_(rowsAntes[i].slice()));
  }

  for (let i = 0; i < rowsDesglosar.length; i++) {
    const row = rowsDesglosar[i];
    const legajo = normalizarLegajo_(row[0]);
    const item = acumulado[legajo];

    if (!legajo || !item || !item.registrosDesglosados || !item.registrosDesglosados.length) {
      salida.push(limpiarColumnasDesglose_(row.slice()));
      continue;
    }

    const registroDesglosado = item.registrosDesglosados.shift();
    const filasRegistro = expandirFilasDesglosadas_(registroDesglosado, item.tope);
    for (let j = 0; j < filasRegistro.length; j++) {
      salida.push(filasRegistro[j]);
    }
  }

  for (let i = 0; i < rowsDespues.length; i++) {
    salida.push(limpiarColumnasDesglose_(rowsDespues[i].slice()));
  }

  return salida;
}

function expandirFilasDesglosadas_(registroDesglosado, tope) {
  const rowBase = limpiarColumnasDesglose_(registroDesglosado.registro.row.slice());
  const filas = [];
  const tramos = [];
  const horasSolicitadas = redondear_(numero_(registroDesglosado.horasSolicitadas));
  const horasAsignadas = redondear_(
    numero_(registroDesglosado.extra100) +
    numero_(registroDesglosado.normal) +
    numero_(registroDesglosado.extra50)
  );
  const estadoPendiente = horasAsignadas > 0 ? 'Parcialmente denegado' : 'Denegado';

  if (numero_(registroDesglosado.extra100) > 0) {
    tramos.push({ tipo: 'Hora al 100%', horas: registroDesglosado.extra100 });
  }

  if (numero_(registroDesglosado.normal) > 0) {
    tramos.push({ tipo: 'Hora normal', horas: registroDesglosado.normal });
  }

  if (numero_(registroDesglosado.extra50) > 0) {
    tramos.push({ tipo: 'Hora al 50%', horas: registroDesglosado.extra50 });
  }

  if (tramos.length <= 1) {
    const fila = rowBase.slice();
    const asignadas = tramos.length ? redondear_(tramos[0].horas) : 0;
    fila[12] = tramos.length ? tramos[0].tipo : estadoPendiente;
    fila[13] = redondear_(tope);
    fila[18] = asignadas || '';
    fila[19] = redondear_(horasSolicitadas - asignadas);
    filas.push(fila);
    return filas;
  }

  for (let i = 0; i < tramos.length; i++) {
    agregarFilaTipo_(filas, rowBase, tramos[i].tipo, tramos[i].horas, tope);
  }

  if (registroDesglosado.diferencia > 0) {
    const filaPendiente = rowBase.slice();
    filaPendiente[12] = estadoPendiente;
    filaPendiente[13] = redondear_(tope);
    filaPendiente[14] = redondear_(registroDesglosado.diferencia);
    filaPendiente[18] = '';
    filaPendiente[19] = redondear_(registroDesglosado.diferencia);
    filas.push(filaPendiente);
  }

  return filas;
}

function agregarFilaTipo_(filas, rowBase, tipo, horas, tope) {
  const cantidad = redondear_(Math.max(0, numero_(horas)));
  if (!cantidad) return;

  const fila = rowBase.slice();
  fila[12] = tipo;
  fila[13] = redondear_(tope);
  fila[14] = cantidad;
  fila[18] = cantidad;
  fila[19] = 0;
  filas.push(fila);
}

function limpiarColumnasDesglose_(row) {
  row[12] = '';
  row[13] = '';
  row[18] = '';
  row[19] = '';
  return row;
}

function compactarFilasOperativas_(rows) {
  const compactadas = [];
  let actual = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].slice();
    const legajo = normalizarLegajo_(row[0]);

    if (!legajo) {
      if (actual) {
        compactadas.push(finalizarFilaCompactada_(actual));
        actual = null;
      }
      compactadas.push(row);
      continue;
    }

    const clave = construirClaveFilaOperativa_(row);
    const horasOriginales = obtenerHorasOriginalesFilaOperativa_(row);

    if (!actual || actual.clave !== clave) {
      if (actual) {
        compactadas.push(finalizarFilaCompactada_(actual));
      }

      actual = {
        clave,
        row: limpiarColumnasDesglose_(row.slice()),
        horas: 0,
      };
    }

    actual.horas = redondear_(actual.horas + horasOriginales);
  }

  if (actual) {
    compactadas.push(finalizarFilaCompactada_(actual));
  }

  return compactadas;
}

function finalizarFilaCompactada_(item) {
  const row = item.row.slice();
  row[14] = redondear_(item.horas);
  return row;
}

function construirClaveFilaOperativa_(row) {
  const partes = [];

  for (let i = 0; i < row.length; i++) {
    if (i === 12 || i === 13 || i === 14 || i === 18 || i === 19) continue;
    partes.push(normalizarTexto_(row[i]));
  }

  return partes.join('|');
}

function obtenerHorasOriginalesFilaOperativa_(row) {
  const solicitadas = numero_(row[14]);
  const asignadas = numero_(row[18]);
  const diferencia = numero_(row[19]);

  return redondear_(Math.max(solicitadas, asignadas + diferencia));
}

function esPedidoEvento_(valor) {
  return normalizarTexto_(valor) === 'si';
}

function escribirDesgloseEnOperativa_(hoja, filas, novedadesPorLegajo) {
  hoja.clearContents();

  const range = hoja.getRange(1, 1, filas.length, filas[0].length);
  range.setValues(filas);

  const fondos = [];
  const colorNovedad = '#f4c022';
  const colorDenegado = '#f4cccc';
  const colorEvento = '#cfe2f3';

  for (let i = 0; i < filas.length; i++) {
    const row = filas[i];
    const legajo = normalizarLegajo_(row[0]);
    const fecha = parseFecha_(row[10]);
    const fechaYmd = fecha ? ymd_(fecha) : '';
    const estadoPedido = normalizarTexto_(row[12]);
    const esEvento = esPedidoEvento_(row[20]);
    const fechasNovedad = (novedadesPorLegajo && novedadesPorLegajo[legajo]) || [];
    const tieneNovedad = !!(legajo && fechaYmd && fechasNovedad.indexOf(fechaYmd) >= 0);
    const estaDenegado = estadoPedido === 'denegado' || estadoPedido === 'parcialmente denegado';
    const color = esEvento ? colorEvento : (estaDenegado ? colorDenegado : (tieneNovedad ? colorNovedad : null));

    fondos.push(new Array(filas[0].length).fill(color));
  }

  range.setBackgrounds(fondos);
}

function contarLegajosEnFilas_(rows) {
  let total = 0;

  for (let i = 0; i < rows.length; i++) {
    if (normalizarLegajo_(rows[i][0])) total++;
  }

  return total;
}

function construirMensajeEjecucion_(resultado, datos) {
  return [
    `${resultado.filas.length} legajos procesados`,
    `filas operativas: ${resultado.diagnostico.filasConsideradas}`,
    `rango usado: ${resultado.diagnostico.desdeFilaUsada}-${resultado.diagnostico.hastaFilaUsada}`,
    `modo: ${resultado.diagnostico.modoRango}`,
    `horas reales externas: ${Math.max(0, datos.horasReales.length - 1)}`,
  ].join(' | ');
}

function escribirResultados_(hoja, filas) {
  hoja.clearContents();
  hoja.getRange(1, 1, 1, CFG.RESULT_HEADERS.length).setValues([CFG.RESULT_HEADERS]);

  if (filas.length) {
    hoja.getRange(2, 1, filas.length, CFG.RESULT_HEADERS.length).setValues(filas);
  }
}

function agruparNovedadesPorLegajo_(data, ctx) {
  const out = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const legajo = normalizarLegajo_(row[0]);
    if (!legajo) continue;

    const desde = parseFecha_(row[2]);
    const hasta = parseFecha_(row[3] || row[2]);
    if (!desde || !hasta) continue;

    const fechas = expandirRango_(desde, hasta)
      .filter((f) => f.getFullYear() === ctx.anio && (f.getMonth() + 1) === ctx.mes)
      .map((f) => ymd_(f));

    if (!out[legajo]) out[legajo] = [];
    out[legajo].push.apply(out[legajo], fechas);
  }

  return out;
}

function indexarNominaPorLegajo_(data) {
  const indices = resolverIndicesNomina_(data[0] || []);
  const out = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const legajo = normalizarLegajo_(row[indices.legajo]);
    if (!legajo) continue;

    out[legajo] = {
      apellido: valorEnIndice_(row, indices.apellido),
      nombre: valorEnIndice_(row, indices.nombre),
      jornada: valorEnIndice_(row, indices.jornada),
      categoria: valorEnIndice_(row, indices.categoria),
      servicio: valorEnIndice_(row, indices.servicio),
    };
  }

  return out;
}

function resumirHorasRealesPorLegajo_(data) {
  const out = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const legajo = normalizarLegajo_(row[1]);
    if (!legajo) continue;

    out[legajo] = redondear_(numero_(out[legajo]) + numero_(row[2]));
  }

  return out;
}

function indexarValoresHora_(data) {
  const out = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // A = Convenio
    // B = Categoria
    // C = Tipo
    // D = Valor
    const categoria = normalizarTexto_(row[1]);
    const tipo = normalizarTexto_(row[2]);
    const valor = numero_(row[3]);

    if (!categoria || !tipo) continue;

    if (!out[categoria]) {
      out[categoria] = {
        normal: 0,
        extra50: 0,
        extra100: 0,
      };
    }

    if (tipo === 'normal') {
      out[categoria].normal = valor;
    } else if (tipo === '50%' || tipo === '50') {
      out[categoria].extra50 = valor;
    } else if (tipo === '100%' || tipo === '100') {
      out[categoria].extra100 = valor;
    }
  }

  return out;
}

function cargarMapaTiposServicio_() {
  const archivo = SpreadsheetApp.openByUrl(CFG.SERVICE_CATALOG_URL);
  const hoja = archivo.getSheetByName(CFG.SERVICE_CATALOG_SHEET);

  if (!hoja) {
    throw new Error(`No existe la hoja ${CFG.SERVICE_CATALOG_SHEET} en el archivo de servicios`);
  }

  const data = hoja.getDataRange().getValues();
  const mapa = {
    exacto: {},
    simplificado: {},
    entradas: [],
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nombreServicio = String(row[CFG.SERVICE_CATALOG_SERVICE_COL - 1] || '').trim();
    const tipoServicio = normalizarTipoServicio_(row[CFG.SERVICE_CATALOG_TYPE_COL - 1]);

    if (!nombreServicio || !tipoServicio) continue;

    const nombreNormalizado = normalizarTexto_(nombreServicio);
    const nombreSimplificado = simplificarTextoComparacion_(nombreServicio);

    mapa.exacto[nombreNormalizado] = tipoServicio;
    if (nombreSimplificado) {
      mapa.simplificado[nombreSimplificado] = tipoServicio;
    }
    mapa.entradas.push({
      original: nombreServicio,
      normalizado: nombreNormalizado,
      simplificado: nombreSimplificado,
      tipo: tipoServicio,
    });
  }

  return mapa;
}

function resolverTipoServicio_(servicioOrigen, servicioAlternativo, tiposServicio) {
  const candidatos = [servicioOrigen];

  if (!String(servicioOrigen || '').trim() && String(servicioAlternativo || '').trim()) {
    candidatos.push(servicioAlternativo);
  }

  for (let i = 0; i < candidatos.length; i++) {
    const texto = String(candidatos[i] || '').trim();
    const key = normalizarTexto_(texto);
    const simplificado = simplificarTextoComparacion_(texto);
    if (!key && !simplificado) continue;

    const tipo = resolverTipoDesdeCatalogo_(key, simplificado, tiposServicio);
    if (tipo) return tipo;
  }

  return '';
}

function resolverTipoDesdeCatalogo_(key, simplificado, tiposServicio) {
  if (key && tiposServicio.exacto[key]) {
    return tiposServicio.exacto[key];
  }

  if (simplificado && tiposServicio.simplificado[simplificado]) {
    return tiposServicio.simplificado[simplificado];
  }

  if (!simplificado) return '';

  for (let i = 0; i < tiposServicio.entradas.length; i++) {
    const entrada = tiposServicio.entradas[i];
    if (!entrada.simplificado) continue;

    if (simplificado.indexOf(entrada.simplificado) >= 0 || entrada.simplificado.indexOf(simplificado) >= 0) {
      return entrada.tipo;
    }
  }

  return '';
}

function normalizarTipoServicio_(tipoServicio) {
  const key = normalizarTexto_(tipoServicio);

  if (!key) return '';
  if (key === 'colegio') return 'Colegio';
  if (key === 'hospital') return 'Hospital';
  if (key === 'supermercado') return 'Supermercado';
  if (key === 'lunes a sabado' || key === 'lunes a sábado') return 'Lunes a Sábado';

  return String(tipoServicio || '').trim();
}

function simplificarTextoComparacion_(txt) {
  return normalizarTexto_(txt).replace(/[^a-z0-9]/g, '');
}

function calcularTopeMensual_(hsTeoricasBase, hsTeoricasFinales) {
  const reduccion = Math.max(0, numero_(hsTeoricasBase) - numero_(hsTeoricasFinales));
  return redondear_(Math.max(0, CFG.TOPE_BASE_MENSUAL - reduccion));
}

function calcularHorasTeoricasBaseMensuales_(servicio, jornadaSemanal, mes, anio) {
  const diasLaborales = obtenerDiasLaborales_(servicio, anio, mes, new Set());
  let diasFinales = diasLaborales.slice();
  const jornada = Number(jornadaSemanal || 0);

  if (servicio === 'Supermercado' || servicio === 'Hospital') {
    const francos = jornada === 44 ? 6 : 8;
    diasFinales = diasFinales.slice(0, Math.max(0, diasFinales.length - francos));
  }

  let horas = 0;

  diasFinales.forEach((dia) => {
    if (servicio === 'Supermercado' && jornada === 24) {
      horas += 4;
      return;
    }

    if (servicio === 'Supermercado' && jornada === 36) {
      horas += 6;
      return;
    }

    if (servicio === 'Lunes a Sábado' && jornada === 44) {
      horas += dia.getDay() === 6 ? 4 : 8;
      return;
    }

    if (servicio === 'Lunes a Sábado' && jornada === 36) {
      horas += 6;
      return;
    }

    if (jornada === 40 || jornada === 44) {
      horas += 8;
    } else if (jornada === 30) {
      horas += 6;
    } else {
      horas += jornada / 5;
    }
  });

  return redondear_(horas);
}

function calcularHorasTeoricasMensuales_(servicio, servicioNombre, jornadaSemanal, mes, anio, ausenciasYmd, feriadosConfig) {
  const ausenciasSet = new Set((ausenciasYmd || []).filter(Boolean));
  const feriados = obtenerFeriadosFinales(servicio, servicioNombre, feriadosConfig);
  const diasLaborales = obtenerDiasLaborales_(servicio, anio, mes, feriados);
  const diasSinAusencias = diasLaborales.filter((d) => !ausenciasSet.has(ymd_(d)));
  const jornada = Number(jornadaSemanal || 0);

  let diasFinales = diasSinAusencias.slice();

  if (servicio === 'Supermercado' || servicio === 'Hospital') {
    const francos = jornada === 44 ? 6 : 8;
    diasFinales = diasFinales.slice(0, Math.max(0, diasFinales.length - francos));
  }

  let horas = 0;

  diasFinales.forEach((dia) => {
    if (servicio === 'Supermercado' && jornada === 24) {
      horas += 4;
      return;
    }

    if (servicio === 'Supermercado' && jornada === 36) {
      horas += 6;
      return;
    }

    if (servicio === 'Lunes a Sábado' && jornada === 44) {
      horas += dia.getDay() === 6 ? 4 : 8;
      return;
    }

    if (servicio === 'Lunes a Sábado' && jornada === 36) {
      horas += 6;
      return;
    }

    if (jornada === 40 || jornada === 44) {
      horas += 8;
    } else if (jornada === 30) {
      horas += 6;
    } else {
      horas += jornada / 5;
    }
  });

  return redondear_(horas);
}

function obtenerFeriadosPorTipo(servicioTipo, feriadosConfig) {
  if (!feriadosConfig || !servicioTipo) return new Set();

  const tipo = normalizarTipoServicio_(servicioTipo);
  return new Set(Array.from((feriadosConfig.porTipo && feriadosConfig.porTipo[tipo]) || []));
}

function obtenerFeriadosPorServicio(servicioNombre, feriadosConfig) {
  if (!feriadosConfig || !servicioNombre) return new Set();

  const servicio = normalizarTexto_(servicioNombre);
  const servicioSimplificado = simplificarTextoComparacion_(servicioNombre);
  const resultado = new Set();
  const catalogo = feriadosConfig.porServicio || {};

  const exactos = catalogo.exacto && catalogo.exacto[servicio];
  if (exactos) {
    exactos.forEach((fecha) => resultado.add(fecha));
  }

  const simplificados = catalogo.simplificado && catalogo.simplificado[servicioSimplificado];
  if (simplificados) {
    simplificados.forEach((fecha) => resultado.add(fecha));
  }

  const entradas = catalogo.entradas || [];
  for (let i = 0; i < entradas.length; i++) {
    const entrada = entradas[i];
    if (!entrada.simplificado || !servicioSimplificado) continue;

    if (
      servicioSimplificado.indexOf(entrada.simplificado) >= 0 ||
      entrada.simplificado.indexOf(servicioSimplificado) >= 0
    ) {
      resultado.add(entrada.fecha);
    }
  }

  return resultado;
}

function obtenerFeriadosFinales(servicioTipo, servicioNombre, feriadosConfig) {
  const feriados = new Set();
  const feriadosTipo = obtenerFeriadosPorTipo(servicioTipo, feriadosConfig);
  const feriadosServicio = obtenerFeriadosPorServicio(servicioNombre, feriadosConfig);

  feriadosTipo.forEach((fecha) => feriados.add(fecha));
  feriadosServicio.forEach((fecha) => feriados.add(fecha));

  return feriados;
}

function resumirConfiguracionFeriados_(feriadosConfig) {
  const tipos = Object.keys((feriadosConfig && feriadosConfig.porTipo) || {}).length;
  const servicios = Object.keys((((feriadosConfig && feriadosConfig.porServicio) || {}).exacto) || {}).length;
  return `tipos con feriados: ${tipos} | servicios con feriados extra: ${servicios}`;
}

function obtenerDiasLaborales_(servicio, anio, mes, feriados) {
  const dias = [];
  const ultimoDia = new Date(anio, mes, 0).getDate();
  const feriadosSet = feriados instanceof Set ? feriados : new Set(feriados || []);

  for (let d = 1; d <= ultimoDia; d++) {
    const dia = new Date(anio, mes - 1, d);
    const clave = ymd_(dia);
    const weekday = dia.getDay();

    if (servicio === 'Supermercado') {
      if (!feriadosSet.has(clave)) dias.push(dia);
      continue;
    }

    if (servicio === 'Hospital') {
      if (!feriadosSet.has(clave)) dias.push(dia);
      continue;
    }

    if (servicio === 'Colegio') {
      if (weekday >= 1 && weekday <= 5 && !feriadosSet.has(clave)) dias.push(dia);
      continue;
    }

    if (servicio === 'Lunes a Sábado') {
      if (weekday >= 1 && weekday <= 6 && !feriadosSet.has(clave)) dias.push(dia);
      continue;
    }

    if (weekday >= 1 && weekday <= 5) {
      dias.push(dia);
    }
  }

  return dias;
}

function desglosarHorasExtra_(hsExtra) {
  return desglosarHorasLegajo_({
    registros: [{ horas: hsExtra }],
    totalDisponible: hsExtra,
    cupoNormal: hsExtra,
  });
}

function resolverIndicesNomina_(headers) {
  const mapa = construirMapaColumnas_(headers);

  return {
    legajo: resolverIndiceColumna_(mapa, ['leg', 'legajo'], 0),
    apellido: resolverIndiceColumna_(mapa, ['apellido', 'apellidos'], -1),
    nombre: resolverIndiceColumna_(mapa, ['nombre', 'nombres'], -1),
    jornada: resolverIndiceColumna_(mapa, ['jornada'], -1),
    categoria: resolverIndiceColumna_(mapa, ['categoria'], -1),
    servicio: resolverIndiceColumna_(mapa, ['servicio origen', 'servicio', 'servicio destino'], -1),
  };
}

function resolverIndicesHorasReales_(headers) {
  const mapa = construirMapaColumnas_(headers);

  return {
    legajo: resolverIndiceColumna_(mapa, ['leg', 'legajo'], -1),
    hsReales: resolverIndiceColumna_(
      mapa,
      ['hs reales', 'horas reales', 'horas trabajadas', 'total horas', 'total hs'],
      -1
    ),
  };
}

function construirMapaColumnas_(headers) {
  const mapa = {};

  for (let i = 0; i < headers.length; i++) {
    const key = normalizarTexto_(headers[i]);
    if (key && mapa[key] === undefined) {
      mapa[key] = i;
    }
  }

  return mapa;
}

function resolverIndiceColumna_(mapa, aliases, fallback) {
  for (let i = 0; i < aliases.length; i++) {
    const key = normalizarTexto_(aliases[i]);
    if (mapa[key] !== undefined) {
      return mapa[key];
    }
  }

  return fallback;
}

function valorEnIndice_(row, index) {
  return index >= 0 ? row[index] || '' : '';
}

function usuarioAutorizadoParaEjecutar_() {
  const usuario = normalizarTexto_(Session.getActiveUser().getEmail());
  if (!usuario) return false;

  const permitidos = CFG.AUTHORIZED_EXECUTORS || [];

  for (let i = 0; i < permitidos.length; i++) {
    if (usuario === normalizarTexto_(permitidos[i])) {
      return true;
    }
  }

  return false;
}

function parseFecha_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return new Date(v);
  }

  const s = String(v || '').trim();
  if (!s) return null;

  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) {
    const partes = s.split('-').map(Number);
    return new Date(partes[0], partes[1] - 1, partes[2]);
  }

  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? Number(`20${m[3]}`) : Number(m[3]);
    return new Date(y, Number(m[2]) - 1, Number(m[1]));
  }

  const n = Number(s);
  if (!isNaN(n)) {
    const d = new Date(Math.round((n - 25569) * 86400 * 1000));
    d.setHours(0, 0, 0, 0);
    return d;
  }

  return null;
}

function parseFechasMultiples_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return [new Date(v)];
  }

  const texto = String(v || '').trim();
  if (!texto) return [];

  const partes = texto.split(/[;,\n]+/);
  const fechas = [];

  for (let i = 0; i < partes.length; i++) {
    const parte = partes[i].trim();
    if (!parte) continue;

    const fecha = parseFecha_(parte);
    if (fecha) fechas.push(fecha);
  }

  if (fechas.length) return fechas;

  const fechaUnica = parseFecha_(v);
  return fechaUnica ? [fechaUnica] : [];
}

function expandirRango_(desde, hasta) {
  const out = [];
  const current = new Date(desde);
  const end = new Date(hasta);
  current.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  while (current <= end) {
    out.push(new Date(current));
    current.setDate(current.getDate() + 1);
  }

  return out;
}

function ymd_(date) {
  return Utilities.formatDate(date, CFG.TZ, 'yyyy-MM-dd');
}

function nombreMes_(mes) {
  return [
    '', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ][Number(mes)] || '';
}

function normalizarTexto_(txt) {
  return String(txt || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function numero_(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const n = Number(String(v || '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function normalizarLegajo_(v) {
  const s = String(v || '').trim();
  return s || '';
}

function redondear_(n) {
  return Math.round(numero_(n) * 100) / 100;
}

function setEstado_(estado, mensaje) {
  const hojaConfig = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_CONFIG);
  if (!hojaConfig) return;

  const posiciones = buscarFilasConfiguracion_(hojaConfig);

  if (posiciones.ultimaEjecucion) {
    hojaConfig.getRange(posiciones.ultimaEjecucion, 2).setValue(new Date());
  } else {
    hojaConfig.getRange(CFG.CONFIG_CELLS.ultimaEjecucion).setValue(new Date());
  }

  if (posiciones.estado) {
    hojaConfig.getRange(posiciones.estado, 2).setValue(estado);
  } else {
    hojaConfig.getRange(CFG.CONFIG_CELLS.estado).setValue(estado);
  }

  if (posiciones.mensaje) {
    hojaConfig.getRange(posiciones.mensaje, 2).setValue(mensaje);
  } else {
    hojaConfig.getRange(CFG.CONFIG_CELLS.mensaje).setValue(mensaje);
  }
}

function buscarFilasConfiguracion_(hojaConfig) {
  const lastRow = hojaConfig.getLastRow();
  const posiciones = {
    ultimaEjecucion: 0,
    estado: 0,
    mensaje: 0,
  };

  if (lastRow <= 0) return posiciones;

  const etiquetas = hojaConfig.getRange(1, 1, lastRow, 1).getValues();

  for (let i = 0; i < etiquetas.length; i++) {
    const etiqueta = normalizarTexto_(etiquetas[i][0]);
    if (etiqueta === 'ultima ejecucion') posiciones.ultimaEjecucion = i + 1;
    if (etiqueta === 'estado') posiciones.estado = i + 1;
    if (etiqueta === 'mensaje') posiciones.mensaje = i + 1;
  }

  return posiciones;
}