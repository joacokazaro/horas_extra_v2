function desglosarHorasLegajo_(params) {
	const servicio = params && params.servicio ? params.servicio : '';
	const servicioNombre = params && params.servicioNombre ? params.servicioNombre : '';
	const registros = (params && params.registros) || [];
	const totalDisponible = Math.max(0, numero_(params && params.totalDisponible));
	const cupoNormal = Math.max(0, numero_(params && params.cupoNormal));
	const feriados = obtenerFeriadosFinales(servicio, servicioNombre, (params && params.feriadosConfig) || null);
	const preparados = [];

	for (let i = 0; i < registros.length; i++) {
		const registro = registros[i] || {};
		const horasSolicitadas = redondear_(Math.max(0, numero_(registro.horas)));
		if (!horasSolicitadas) continue;

		preparados.push({
			registro: registro,
			horasSolicitadas: horasSolicitadas,
			categoriaBase: clasificarHoraExtraBase_(servicio, registro, feriados),
			normal: 0,
			extra50: 0,
			extra100: 0,
			diferencia: 0,
		});
	}

	let restanteDisponible = totalDisponible;

	for (let i = 0; i < preparados.length; i++) {
		const item = preparados[i];
		if (item.categoriaBase !== 'extra100' || restanteDisponible <= 0) continue;

		const asignadas = Math.min(item.horasSolicitadas, restanteDisponible);
		item.extra100 = redondear_(asignadas);
		restanteDisponible = redondear_(restanteDisponible - asignadas);
	}

	const horas100 = sumarHorasPorCampo_(preparados, 'extra100');
	const totalNo100Solicitado = sumarHorasNo100Solicitadas_(preparados);
	const cupoNormalRestante = Math.max(0, Math.min(cupoNormal, totalDisponible) - horas100);
	const pagableNo100 = Math.min(totalNo100Solicitado, restanteDisponible);
	let restante50 = Math.max(0, pagableNo100 - cupoNormalRestante);
	let restanteNormal = Math.max(0, pagableNo100 - restante50);

	for (let i = 0; i < preparados.length; i++) {
		const item = preparados[i];
		if (item.categoriaBase === 'extra100' || restanteDisponible <= 0) continue;

		const pagables = Math.min(item.horasSolicitadas, restanteDisponible);
		const al50 = Math.min(pagables, restante50);
		const normales = Math.min(pagables - al50, restanteNormal);

		item.extra50 = redondear_(al50);
		item.normal = redondear_(normales);

		restante50 = redondear_(Math.max(0, restante50 - al50));
		restanteNormal = redondear_(Math.max(0, restanteNormal - normales));
		restanteDisponible = redondear_(restanteDisponible - pagables);
	}

	for (let i = 0; i < preparados.length; i++) {
		const item = preparados[i];
		const pagadas = numero_(item.normal) + numero_(item.extra50) + numero_(item.extra100);
		item.diferencia = redondear_(Math.max(0, item.horasSolicitadas - pagadas));
	}

	const normal = sumarHorasPorCampo_(preparados, 'normal');
	const extra50 = sumarHorasPorCampo_(preparados, 'extra50');
	const extra100 = sumarHorasPorCampo_(preparados, 'extra100');

	return {
		normal: normal,
		extra50: extra50,
		extra100: extra100,
		total: redondear_(normal + extra50 + extra100),
		registros: preparados,
	};
}

function sumarHorasNo100Solicitadas_(items) {
	let total = 0;

	for (let i = 0; i < items.length; i++) {
		const item = items[i];
		if (!item || item.categoriaBase === 'extra100') continue;
		total += numero_(item.horasSolicitadas);
	}

	return redondear_(total);
}

function clasificarHoraExtraBase_(servicio, registro, feriados) {
	const diaSemana = normalizarTexto_(registro && registro.diaSemana);
	const fecha = parseFecha_(registro && registro.fecha);
	const esSabado = esSabadoExtra_(diaSemana, fecha);
	const esDomingo = esDomingoExtra_(diaSemana, fecha);
	const esFeriado = esFeriadoExtra_(diaSemana, fecha, feriados);
	const esSabadoDespues13 = esSabado && !esHastaSabado13_(diaSemana);

	if (servicio === 'Colegio') {
		if (esFeriado || esDomingo || esSabado) return 'extra100';
		return 'normal';
	}

	if (servicio === 'Lunes a Sábado') {
		if (esFeriado || esDomingo || esSabadoDespues13) return 'extra100';
		return 'normal';
	}

	if (servicio === 'Hospital' || servicio === 'Supermercado') {
		return 'normal';
	}

	return 'normal';
}

function calcularValorDesglose_(desglose, valoresCategoria) {
	const valores = valoresCategoria || {};
	return redondear_(
		numero_(desglose && desglose.normal) * numero_(valores.normal) +
		numero_(desglose && desglose.extra50) * numero_(valores.extra50) +
		numero_(desglose && desglose.extra100) * numero_(valores.extra100)
	);
}

function esSabadoExtra_(diaSemana, fecha) {
	if (esEtiquetaSabadoExplicita_(diaSemana)) return true;
	return !!(fecha && fecha.getDay() === 6);
}

function esDomingoExtra_(diaSemana, fecha) {
	if (diaSemana && diaSemana.indexOf('domingo') >= 0) return true;
	return !!(fecha && fecha.getDay() === 0);
}

function esFeriadoExtra_(diaSemana, fecha, feriados) {
	if (diaSemana && diaSemana.indexOf('feriado') >= 0) return true;
	if (!fecha || !(feriados instanceof Set)) return false;
	return feriados.has(ymd_(fecha));
}

function esHastaSabado13_(diaSemana) {
	if (!diaSemana) return false;

	return (
		diaSemana.indexOf('hasta las 13') >= 0 ||
		diaSemana.indexOf('hasta 13') >= 0 ||
		diaSemana.indexOf('lunes a sabado') >= 0
	);
}

function esEtiquetaSabadoExplicita_(diaSemana) {
	if (!diaSemana) return false;

	return (
		diaSemana === 'sabado' ||
		diaSemana.indexOf('sabado despues de las 13') >= 0 ||
		diaSemana.indexOf('sabado después de las 13') >= 0 ||
		diaSemana.indexOf('sabado despues de 13') >= 0 ||
		diaSemana.indexOf('sabado después de 13') >= 0
	);
}

function sumarHorasPorCampo_(items, campo) {
	let total = 0;

	for (let i = 0; i < items.length; i++) {
		total += numero_(items[i] && items[i][campo]);
	}

	return redondear_(total);
}
