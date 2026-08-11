/**
 * BACKEND — Gestión Institucional La Eugenia
 * Módulo: Autoridades
 *
 * INSTRUCCIONES DE DESPLIEGUE (hace falta tu cuenta de Google, no lo puedo hacer yo):
 * 1. Crear un Google Sheet nuevo. Crear en él una hoja llamada exactamente "AUTORIDADES"
 *    con estas columnas en la fila 1:
 *    ID_AUTORIDAD | ORGANO | CARGO | GRUPO_ESTATUTARIO | NOMBRE | DNI | FECHA_INICIO_MANDATO | FECHA_FIN_MANDATO | ESTADO | ID_SOCIO
 * 2. Extensiones → Apps Script. Pegar este código completo, reemplazando el contenido default.
 * 3. Reemplazar SHEET_ID abajo con el ID de tu Sheet nuevo (está en la URL).
 * 4. Implementar → Nueva implementación → Aplicación web → Ejecutar como "yo" → Acceso "cualquiera".
 * 5. Copiar la URL /exec que te da y pegarla en index.html, en EXEC_URL_GI.
 */

const SHEET_ID = '1mT5ezYh7R-59APw6JF2B2O6EprTj7xbguys1lpvh6dw';
const SHEET_ID_BOLETIN = '1EplYcC36NhlBUqwizwS_YdeB2LFEFuRjXsj5Q_kKQLw';
const HOJA_AUTORIDADES = 'AUTORIDADES';

const PINES_CGV_SHEET_ID = '1altioUeYlQW8NXWYVjHf5v_4ocJuQ-K9N0EfdYQSoBc';

function debugPin(pin) {
  const info = { pinRecibido: pin };
  try {
    const ss = SpreadsheetApp.openById(PINES_CGV_SHEET_ID);
    info.nombresDeHojas = ss.getSheets().map(s => s.getName());

    const hoja = ss.getSheetByName('PINES_CGV');
    info.encontroHojaPorNombre = !!hoja;

    const hojaUsada = hoja || ss.getSheets()[0];
    const datos = hojaUsada.getDataRange().getValues();
    info.totalFilas = datos.length;
    info.primeras3Filas = datos.slice(0, 3);

    let filaHeaders = -1;
    for (let i = 0; i < datos.length; i++) {
      if (datos[i].indexOf('PIN') !== -1) { filaHeaders = i; break; }
    }
    info.filaHeadersDetectada = filaHeaders;
    if (filaHeaders !== -1) info.headersEncontrados = datos[filaHeaders];

    info.resultadoValidacion = validarPinInterno(pin, 'gestion-institucional');
  } catch (err) {
    info.excepcion = err.message;
  }
  return { ok: true, debug: info };
}

function validarPinInterno(pin, moduloRequerido) {
  if (!pin) return false;
  const hoja = SpreadsheetApp.openById(PINES_CGV_SHEET_ID).getSheetByName('PINES_CGV');
  if (!hoja) return false;
  const datos = hoja.getDataRange().getValues();

  let filaHeaders = -1;
  for (let i = 0; i < datos.length; i++) {
    if (datos[i].indexOf('PIN') !== -1) { filaHeaders = i; break; }
  }
  if (filaHeaders === -1) return false;

  const headers = datos[filaHeaders];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = filaHeaders + 1; i < datos.length; i++) {
    const fila = datos[i];
    if (String(fila[idx.PIN]) === String(pin)) {
      if (fila[idx.ACTIVO] !== 'SI') return false;
      const modulos = String(fila[idx.MODULOS] || '');
      return modulos === 'ALL' || modulos.split(',').map(m => m.trim()).includes(moduloRequerido);
    }
  }
  return false;
}

// Acciones que cuestan dinero o escriben datos: requieren PIN válido verificado en el servidor,
// no solo en la pantalla. El resto (listar/consultar) queda sin este requisito por ahora.
const ACCIONES_PROTEGIDAS = ['guardarAutoridad', 'guardarNota', 'eliminarNota', 'generarBorradorActa', 'actualizarActa', 'guardarActaHistorica', 'procesarBalance', 'actualizarEstadoBalance', 'cerrarYAbrirNuevoEjercicio', 'actualizarObservacionBalance', 'guardarNovedad', 'actualizarNovedad', 'generarBorradorMemoria', 'registrarInformeRevisor', 'generarConvocatoriaYDocumentos', 'actualizarDocumento', 'guardarNovedadesSeleccionadas', 'extraerNovedadesDeChat', 'eliminarNovedad', 'guardarConfigMemoria', 'sembrarAsambleaReal', 'guardarAsambleaEjercicio', 'corregirFechaAsambleaExistente', 'limpiarDuplicadosAsambleas', 'marcarAsambleaCelebrada', 'registrarAutoridadesElectas', 'generarActaAsamblea'];

function doGet(e) {
  const action = e.parameter.action;

  if (ACCIONES_PROTEGIDAS.indexOf(action) !== -1) {
    if (!validarPinInterno(e.parameter.pin, 'gestion-institucional')) {
      return jsonpResponse({ ok: false, error: 'PIN inválido o sin permiso para este módulo' }, e.parameter.callback);
    }
  }

  let resultado;
  try {
    if (action === 'listarAutoridades') {
      resultado = listarAutoridades();
    } else if (action === 'guardarAutoridad') {
      resultado = guardarAutoridad(e.parameter);
    } else if (action === 'guardarNota') {
      resultado = guardarNota(e.parameter);
    } else if (action === 'eliminarNota') {
      resultado = eliminarNota(e.parameter);
    } else if (action === 'listarNotasPendientes') {
      resultado = listarNotasPendientes();
    } else if (action === 'listarActas') {
      resultado = listarActas();
    } else if (action === 'obtenerActa') {
      resultado = obtenerActa(e.parameter.idActa);
    } else if (action === 'generarBorradorActa') {
      resultado = generarBorradorActa(e.parameter);
    } else if (action === 'actualizarActa') {
      resultado = actualizarActa(e.parameter);
    } else if (action === 'guardarActaHistorica') {
      resultado = guardarActaHistorica(e.parameter);
    } else if (action === 'procesarBalance') {
      resultado = procesarBalance();
    } else if (action === 'listarBalances') {
      resultado = listarBalances();
    } else if (action === 'obtenerBalance') {
      resultado = obtenerBalance(e.parameter.idBalance);
    } else if (action === 'actualizarEstadoBalance') {
      resultado = actualizarEstadoBalance(e.parameter);
    } else if (action === 'debugPin') {
      resultado = debugPin(e.parameter.pin);
    } else if (action === 'version') {
      resultado = { ok: true, version: 'v43-fix-fecha-en-rango-ejercicio-11ago-1430' };
    } else if (action === 'obtenerEjercicioActivo') {
      resultado = obtenerEjercicioActivo();
    } else if (action === 'obtenerResumenDashboard') {
      resultado = obtenerResumenDashboard();
    } else if (action === 'listarEjercicios') {
      resultado = listarEjercicios();
    } else if (action === 'cerrarYAbrirNuevoEjercicio') {
      resultado = cerrarYAbrirNuevoEjercicio(e.parameter);
    } else if (action === 'actualizarObservacionBalance') {
      resultado = actualizarObservacionBalance(e.parameter);
    } else if (action === 'guardarNovedad') {
      resultado = guardarNovedad(e.parameter);
    } else if (action === 'listarNovedades') {
      resultado = listarNovedades(e.parameter);
    } else if (action === 'actualizarNovedad') {
      resultado = actualizarNovedad(e.parameter);
    } else if (action === 'generarBorradorMemoria') {
      resultado = generarBorradorMemoria(e.parameter);
    } else if (action === 'registrarInformeRevisor') {
      resultado = registrarInformeRevisor(e.parameter);
    } else if (action === 'generarConvocatoriaYDocumentos') {
      resultado = generarConvocatoriaYDocumentos(e.parameter);
    } else if (action === 'obtenerAsambleaEjercicio') {
      resultado = obtenerAsambleaEjercicio();
    } else if (action === 'guardarAsambleaEjercicio') {
      resultado = guardarAsambleaEjercicio(e.parameter);
    } else if (action === 'corregirFechaAsambleaExistente') {
      resultado = corregirFechaAsambleaExistente();
    } else if (action === 'limpiarDuplicadosAsambleas') {
      resultado = limpiarDuplicadosAsambleas();
    } else if (action === 'marcarAsambleaCelebrada') {
      resultado = marcarAsambleaCelebrada(e.parameter);
    } else if (action === 'registrarAutoridadesElectas') {
      resultado = registrarAutoridadesElectas(e.parameter);
    } else if (action === 'generarActaAsamblea') {
      resultado = generarActaAsamblea(e.parameter);
    } else if (action === 'diagnosticoAsambleas') {
      resultado = diagnosticoAsambleas();
    } else if (action === 'sembrarAsambleaReal') {
      resultado = sembrarAsambleaReal();
    } else if (action === 'obtenerConfigMemoria') {
      resultado = obtenerConfigMemoria();
    } else if (action === 'guardarConfigMemoria') {
      resultado = guardarConfigMemoria(e.parameter);
    } else if (action === 'obtenerUltimosSocios') {
      resultado = obtenerUltimosSocios(e.parameter);
    } else if (action === 'listarDocumentos') {
      resultado = listarDocumentos(e.parameter);
    } else if (action === 'actualizarDocumento') {
      resultado = actualizarDocumento(e.parameter);
    } else if (action === 'extraerNovedadesDeChat') {
      resultado = extraerNovedadesDeChat(e.parameter);
    } else if (action === 'listarNovedadesDesdeBoletin') {
      resultado = listarNovedadesDesdeBoletin();
    } else if (action === 'guardarNovedadesSeleccionadas') {
      resultado = guardarNovedadesSeleccionadas(e.parameter);
    } else if (action === 'eliminarNovedad') {
      resultado = eliminarNovedad(e.parameter);
    } else {
      resultado = { ok: false, error: 'Acción no reconocida' };
    }
  } catch (err) {
    resultado = { ok: false, error: err.message };
  }
  return jsonpResponse(resultado, e.parameter.callback);
}

function jsonpResponse(obj, callback) {
  const json = JSON.stringify(obj);
  const output = callback ? (callback + '(' + json + ')') : json;
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getHoja() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_AUTORIDADES);
}

function listarAutoridades() {
  const hoja = getHoja();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const filas = datos.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const anioActual = new Date().getFullYear();
  const autoridades = [];
  const vencenEsteAnio = [];

  filas.forEach(fila => {
    if (!fila[idx.NOMBRE]) return; // fila vacía
    const fechaFin = fila[idx.FECHA_FIN_MANDATO];
    const anioFin = fechaFin ? new Date(fechaFin).getFullYear() : null;
    let vence = 'ok';
    if (anioFin === anioActual) {
      vence = 'este_anio';
      vencenEsteAnio.push(fila[idx.CARGO] + ' (' + fila[idx.NOMBRE] + ')');
    } else if (anioFin === anioActual + 1) {
      vence = 'proximo_anio';
    }

    autoridades.push({
      idAutoridad: fila[idx.ID_AUTORIDAD],
      organo: fila[idx.ORGANO],
      cargo: fila[idx.CARGO],
      grupoEstatutario: fila[idx.GRUPO_ESTATUTARIO],
      nombre: fila[idx.NOMBRE],
      dni: fila[idx.DNI],
      fechaInicioMandato: fila[idx.FECHA_INICIO_MANDATO] ? Utilities.formatDate(new Date(fila[idx.FECHA_INICIO_MANDATO]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      fechaFinMandato: fechaFin ? Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      estado: fila[idx.ESTADO],
      vence: vence
    });
  });

  return { ok: true, autoridades: autoridades.filter(a => a.estado === 'VIGENTE'), vencenEsteAnio: vencenEsteAnio };
}

// Grupo estatutario automático por cargo (Art. 21/22 del Estatuto)
const GRUPO_POR_CARGO = {
  'Presidente': 2, 'Secretario': 2, 'Tesorero': 2,
  'Vocal Titular 1°': 2, 'Vocal Titular 2°': 2, 'Vocal Titular 3°': 2,
  'Vicepresidente': 1, 'Prosecretario': 1, 'Protesorero': 1,
  'Vocal Suplente 1°': 1, 'Vocal Suplente 2°': 1, 'Vocal Suplente 3°': 1
};

function calcularFechaFinMandato(organo, cargo, fechaInicio) {
  const inicio = new Date(fechaInicio);
  if (organo === 'REVISORA_CUENTAS') {
    // Renovación anual (Art. 35)
    return new Date(inicio.getFullYear() + 1, inicio.getMonth(), inicio.getDate());
  }
  // Comisión Directiva (Art. 22) y Arquitectura (Art. 37): cada 2 años
  return new Date(inicio.getFullYear() + 2, inicio.getMonth(), inicio.getDate());
}

function guardarAutoridad(params) {
  // TODO: reforzar validación de PIN contra PINES_CGV antes de escribir (mismo patrón que Comunicador)
  const hoja = getHoja();
  const organo = params.organo;
  const cargo = params.cargo;
  const nombre = params.nombre;
  const dni = params.dni || '';
  const fechaInicio = params.fechaInicio;

  // Si ya existe alguien VIGENTE en ese cargo, lo cierra (FINALIZADO) antes de crear el nuevo
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][idx.ORGANO] === organo && datos[i][idx.CARGO] === cargo && datos[i][idx.ESTADO] === 'VIGENTE') {
      hoja.getRange(i + 1, idx.ESTADO + 1).setValue('FINALIZADO');
    }
  }

  const grupo = GRUPO_POR_CARGO[cargo] || '';
  const fechaFin = calcularFechaFinMandato(organo, cargo, fechaInicio);
  const nuevoId = 'AUT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

  hoja.appendRow([nuevoId, organo, cargo, grupo, nombre, dni, new Date(fechaInicio), fechaFin, 'VIGENTE', '']);

  return { ok: true, idAutoridad: nuevoId };
}

// ============ BITÁCORA ============
const HOJA_BITACORA = 'BITACORA';

function getHojaBitacora() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_BITACORA);
}

function guardarNota(params) {
  const hoja = getHojaBitacora();
  const id = 'NOTA-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  hoja.appendRow([id, new Date(), params.cargadoPor || '', params.texto, false, '']);
  return { ok: true, idNota: id };
}

function eliminarNota(params) {
  const hoja = getHojaBitacora();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_NOTA]) === String(params.idNota)) {
      hoja.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Nota no encontrada' };
}

function listarNotasPendientes() {
  const hoja = getHojaBitacora();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const notas = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_NOTA]) continue;
    if (fila[idx.PROCESADA] === true || fila[idx.PROCESADA] === 'TRUE') continue;
    if (String(fila[idx.TEXTO]).indexOf('PENDIENTE EJERCICIO') === 0) continue;
    notas.push({
      idNota: fila[idx.ID_NOTA],
      fechaCarga: fila[idx.FECHA_CARGA] ? Utilities.formatDate(new Date(fila[idx.FECHA_CARGA]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
      cargadoPor: fila[idx.CARGADO_POR],
      texto: fila[idx.TEXTO]
    });
  }
  return { ok: true, notas: notas };
}

// ============ ACTAS ============
const HOJA_ACTAS = 'ACTAS';
const ULTIMA_ACTA_HISTORICA = 562; // última acta real conocida antes de usar este sistema

function getHojaActas() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_ACTAS);
}

function listarActas() {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const actas = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_ACTA]) continue;
    actas.push({
      idActa: fila[idx.ID_ACTA],
      idEjercicio: fila[idx.ID_EJERCICIO] || '(sin asignar)',
      fechaReunion: fila[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(fila[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      estado: fila[idx.ESTADO]
    });
  }
  actas.sort((a, b) => Number(b.idActa) - Number(a.idActa));
  return { ok: true, actas: actas };
}

function obtenerUltimoNumeroActa() {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let maxId = ULTIMA_ACTA_HISTORICA;
  for (let i = 1; i < datos.length; i++) {
    const n = Number(datos[i][idx.ID_ACTA]);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return maxId;
}

function obtenerActa(idActa) {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_ACTA]) === String(idActa)) {
      const fila = datos[i];
      return {
        ok: true,
        acta: {
          idActa: fila[idx.ID_ACTA],
          idActaAnterior: fila[idx.ID_ACTA_ANTERIOR],
          fechaReunion: fila[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(fila[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          horaInicio: formatearHoraSegura(fila[idx.HORA_INICIO]),
          horaFin: formatearHoraSegura(fila[idx.HORA_FIN]),
          presentes: fila[idx.PRESENTES],
          puntos: fila[idx.PUNTOS] ? JSON.parse(fila[idx.PUNTOS]) : [],
          estado: fila[idx.ESTADO]
        }
      };
    }
  }
  return { ok: false, error: 'Acta no encontrada' };
}

function generarBorradorActa(params) {
  const hojaActas = getHojaActas();
  const hojaBitacora = getHojaBitacora();

  const ultimoNumero = obtenerUltimoNumeroActa();
  const nuevoNumero = ultimoNumero + 1;

  // Punto 1 siempre fijo, encadenado al acta anterior
  const puntos = [
    { orden: 1, texto: 'Se da lectura al acta N.º ' + ultimoNumero + '. Se aprueba por unanimidad.' }
  ];

  // Trae notas pendientes de la Bitácora como puntos en borrador (texto crudo, a editar)
  const datosBit = hojaBitacora.getDataRange().getValues();
  const headersBit = datosBit[0];
  const idxBit = {};
  headersBit.forEach((h, i) => idxBit[h] = i);

  const idsNotasUsadas = [];
  for (let i = 1; i < datosBit.length; i++) {
    const fila = datosBit[i];
    if (!fila[idxBit.ID_NOTA]) continue;
    if (fila[idxBit.PROCESADA] === true || fila[idxBit.PROCESADA] === 'TRUE') continue;
    // Las notas marcadas para un Ejercicio futuro no se usan en actas del Ejercicio actual
    if (String(fila[idxBit.TEXTO]).indexOf('PENDIENTE EJERCICIO') === 0) continue;
    puntos.push({ orden: puntos.length + 1, texto: fila[idxBit.TEXTO] });
    idsNotasUsadas.push({ row: i + 1, id: fila[idxBit.ID_NOTA] });
  }

  hojaActas.appendRow([
    nuevoNumero,
    (function() { const e = obtenerEjercicioActivo(); return e.ok ? e.ejercicio.idEjercicio : (params.idEjercicio || ''); })(),
    new Date(params.fechaReunion),
    params.horaInicio || '',
    '',
    params.presentes || '',
    JSON.stringify(puntos),
    'BORRADOR',
    ultimoNumero
  ]);
  hojaActas.getRange(hojaActas.getLastRow(), 4).setNumberFormat('@').setValue(params.horaInicio || ''); // refuerza texto

  // Marca las notas usadas como procesadas
  idsNotasUsadas.forEach(n => {
    hojaBitacora.getRange(n.row, idxBit.PROCESADA + 1).setValue(true);
    hojaBitacora.getRange(n.row, idxBit.ID_ACTA_DESTINO + 1).setValue(nuevoNumero);
  });

  return { ok: true, idActa: nuevoNumero, puntos: puntos };
}

function guardarActaHistorica(params) {
  const hoja = getHojaActas();
  hoja.appendRow([
    params.idActa,
    params.idEjercicio,
    new Date(params.fechaReunion),
    params.horaInicio,
    params.horaFin,
    params.presentes,
    params.puntos,
    params.estado || 'FIRMADA',
    params.idActaAnterior || ''
  ]);
  const fila = hoja.getLastRow();
  hoja.getRange(fila, 4).setNumberFormat('@').setValue(params.horaInicio); // refuerza texto
  hoja.getRange(fila, 5).setNumberFormat('@').setValue(params.horaFin); // refuerza texto
  return { ok: true };
}

function actualizarActa(params) {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_ACTA]) === String(params.idActa)) {
      const row = i + 1;
      if (params.puntos) hoja.getRange(row, idx.PUNTOS + 1).setValue(params.puntos);
      if (params.horaFin) hoja.getRange(row, idx.HORA_FIN + 1).setNumberFormat('@').setValue(params.horaFin);
      if (params.presentes) hoja.getRange(row, idx.PRESENTES + 1).setValue(params.presentes);
      if (params.estado) hoja.getRange(row, idx.ESTADO + 1).setValue(params.estado);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Acta no encontrada' };
}

// ============ BALANCE ============
const HOJA_BALANCES = 'BALANCES';
const CARPETA_BALANCES_ID = '189c_jJ7CkLdHISks0PI2KoBpfC-1_eX1'; // Balances_Gestion_Institucional_La_Eugenia

function getHojaBalances() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_BALANCES);
}

const PROMPT_REVISION_BALANCE = `Sos un asistente contable-legal para una asociación civil sin fines de lucro (Club de Campo "La Eugenia", CUIT 30-63417128-9). Vas a recibir uno o más PDF relacionados con el Balance del ejercicio.

IMPORTANTE — distinción de tipo de documento: puede haber dos tipos de archivo muy distintos, y NUNCA hay que tratarlos como si fueran lo mismo ni compararlos como si usaran la misma base de valuación:
- "ESTADOS_CONTABLES_DEFINITIVOS": el/los documento(s) formal(es) para presentar a la Asamblea (Estado de Situación Patrimonial, Estado de Recursos y Gastos, Estado de Evolución del Patrimonio Neto, Estado de Flujo de Efectivo, Notas, Anexos, Informe del Auditor). Suelen estar en moneda homogénea/ajustada por inflación (RT 54).
- "HOJA_DE_TRABAJO_INTERNA": un balance de sumas y saldos, listado de cuentas contables con números de cuenta (ej. "10000000 ACTIVO"), sin ajuste por inflación, uso interno del estudio contable, NO es lo que se presenta a la Asamblea.

Revisá el/los documento(s) y respondé ÚNICAMENTE con un JSON válido (sin texto adicional, sin markdown, sin backticks), con esta estructura exacta:

{
  "documentosRecibidos": [
    { "tipo": "ESTADOS_CONTABLES_DEFINITIVOS o HOJA_DE_TRABAJO_INTERNA", "descripcion": "breve" }
  ],
  "denominacionEncontrada": "texto",
  "cuitEncontrado": "texto",
  "periodoInicio": "DD/MM/AAAA",
  "periodoCierre": "DD/MM/AAAA",
  "activoTotal": "número o texto (tomado de ESTADOS_CONTABLES_DEFINITIVOS si está disponible; si no, aclarar que es de la hoja de trabajo)",
  "pasivoTotal": "número o texto",
  "patrimonioNeto": "número o texto",
  "superavitEjercicio": "número o texto",
  "cuadraActivoPasivoPN": true o false,
  "observaciones": [
    { "tipo": "CRITICO", "texto": "descripción específica y concreta del hallazgo" },
    { "tipo": "OBSERVACION", "texto": "..." },
    { "tipo": "ADVERTENCIA", "texto": "..." },
    { "tipo": "DATO_CORRECTO", "texto": "..." }
  ]
}

Reglas de revisión (basadas en errores reales ya detectados en este tipo de documento):
1. CRÍTICO: si en cualquier parte del documento (incluyendo el informe del auditor) aparece el nombre de una entidad distinta a "Club de Campo La Eugenia" — son documentos reutilizados de otro cliente sin corregir.
2. CRÍTICO: si los encabezados de fecha de cualquier cuadro comparativo no coinciden con el período informado en la carátula (ej. dice "AL 30/04/2025" pero los valores corresponden al cierre real informado).
3. OBSERVACIÓN: si Activo Total no es exactamente igual a Pasivo Total + Patrimonio Neto, DENTRO DEL MISMO DOCUMENTO (verificar la ecuación contable). No marques como error una diferencia entre la hoja de trabajo y los Estados Contables definitivos — esa diferencia es esperable por el ajuste por inflación, aclaralo como DATO_CORRECTO si corresponde, nunca como error.
4. OBSERVACIÓN: diferencias de centavos entre cuadros que deberían coincidir dentro del mismo documento.
5. ADVERTENCIA: notas contables con redacción ambigua o incompleta (ej. índices de ajuste por inflación sin especificar claramente).
6. Si solo recibiste la HOJA_DE_TRABAJO_INTERNA y no los ESTADOS_CONTABLES_DEFINITIVOS, indicalo como OBSERVACIÓN: "Solo se recibió la hoja de trabajo interna, no los Estados Contables definitivos para presentar a la Asamblea."
7. Nunca corrijas el documento vos mismo. Solo señalá. Si algo requiere juicio profesional, decilo explícitamente en el texto de la observación ("Consultar al profesional responsable").
8. Si no encontrás problemas en una categoría, no incluyas entradas de ese tipo — no inventes observaciones para rellenar.

Respondé solo el JSON.`;

function procesarBalance() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'Falta configurar ANTHROPIC_API_KEY en Script Properties' };
  }

  // Toma TODOS los PDF de la carpeta (un Balance suele venir en más de un archivo)
  const carpeta = DriveApp.getFolderById(CARPETA_BALANCES_ID);
  const archivosIter = carpeta.getFilesByType(MimeType.PDF);
  const archivos = [];
  while (archivosIter.hasNext()) archivos.push(archivosIter.next());

  if (archivos.length === 0) {
    return { ok: false, error: 'No hay ningún PDF en la carpeta de Balances. Subí uno primero.' };
  }

  const contentBlocks = archivos.map(f => ({
    type: 'document',
    source: { type: 'base64', media_type: 'application/pdf', data: Utilities.base64Encode(f.getBlob().getBytes()) }
  }));
  contentBlocks.push({ type: 'text', text: 'Estos son los documentos del Balance (puede ser más de un archivo: Balance General, Estados Contables, etc.). Analizalos en conjunto y respondé solo el JSON según las reglas indicadas.' });

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 16000,
    thinking: { type: 'disabled' },
    system: PROMPT_REVISION_BALANCE,
    messages: [{ role: 'user', content: contentBlocks }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status !== 200) {
    return { ok: false, error: 'Error de la API de Anthropic (' + status + '): ' + response.getContentText() };
  }

  const data = JSON.parse(response.getContentText());

  if (!data.content || !data.content.length) {
    return { ok: false, error: 'La API no devolvió contenido. Respuesta cruda: ' + response.getContentText().substring(0, 500) };
  }

  const bloqueTexto = data.content.find(b => b.type === 'text');
  if (!bloqueTexto || !bloqueTexto.text) {
    return { ok: false, error: 'No se encontró bloque de texto en la respuesta. Tipos recibidos: ' + data.content.map(b => b.type).join(', ') };
  }

  let textoRespuesta = bloqueTexto.text.replace(/```json|```/g, '').trim();

  let analisis;
  try {
    analisis = JSON.parse(textoRespuesta);
  } catch (e) {
    return { ok: false, error: 'La IA no devolvió JSON válido: ' + textoRespuesta.substring(0, 300) };
  }

  // Cada observación arranca como PENDIENTE -- se puede marcar RESUELTA u OMITIDA desde la app, nunca se borra
  (analisis.observaciones || []).forEach(o => { o.estadoObs = 'PENDIENTE'; o.comentario = ''; });

  // Guarda en la hoja BALANCES
  const hoja = getHojaBalances();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let version = 1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][idx.ID_BALANCE]) version++;
  }

  const nuevoId = 'BAL-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const hayCriticos = analisis.observaciones.some(o => o.tipo === 'CRITICO');
  const estadoInicial = hayCriticos ? 'OBSERVADO' : 'EN_REVISION';
  const urls = archivos.map(f => f.getUrl()).join(' | ');

  const ejercicioActivo = obtenerEjercicioActivo();
  const idEjercicioActual = ejercicioActivo.ok ? ejercicioActivo.ejercicio.idEjercicio : '';

  hoja.appendRow([
    nuevoId,
    idEjercicioActual,
    urls,
    version,
    estadoInicial,
    JSON.stringify(analisis),
    JSON.stringify(analisis.observaciones)
  ]);

  // Archiva los PDF ya procesados en una subcarpeta propia, para que la carpeta principal
  // quede siempre vacía y lista para el próximo balance -- así nunca se mezclan períodos.
  const carpetaArchivo = carpeta.createFolder('Procesado_' + nuevoId);
  archivos.forEach(f => {
    carpeta.removeFile(f);
    carpetaArchivo.addFile(f);
  });

  return { ok: true, idBalance: nuevoId, analisis: analisis, estado: estadoInicial, archivosProcesados: archivos.length };
}

function listarBalances() {
  const hoja = getHojaBalances();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const balances = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_BALANCE]) continue;
    balances.push({ idBalance: fila[idx.ID_BALANCE], version: fila[idx.VERSION], estado: fila[idx.ESTADO], idEjercicio: fila[idx.ID_EJERCICIO] || '(sin asignar)' });
  }
  balances.sort((a, b) => b.version - a.version);
  return { ok: true, balances: balances };
}

function obtenerBalance(idBalance) {
  const hoja = getHojaBalances();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_BALANCE]) === String(idBalance)) {
      const fila = datos[i];
      return {
        ok: true,
        balance: {
          idBalance: fila[idx.ID_BALANCE],
          archivoUrl: fila[idx.ARCHIVO_PDF_URL],
          version: fila[idx.VERSION],
          estado: fila[idx.ESTADO],
          datosExtraidos: fila[idx.DATOS_EXTRAIDOS] ? JSON.parse(fila[idx.DATOS_EXTRAIDOS]) : {},
          observaciones: fila[idx.OBSERVACIONES] ? JSON.parse(fila[idx.OBSERVACIONES]) : []
        }
      };
    }
  }
  return { ok: false, error: 'Balance no encontrado' };
}

function actualizarEstadoBalance(params) {
  const hoja = getHojaBalances();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_BALANCE]) === String(params.idBalance)) {
      hoja.getRange(i + 1, idx.ESTADO + 1).setValue(params.estado);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Balance no encontrado' };
}

function actualizarObservacionBalance(params) {
  const hoja = getHojaBalances();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const indice = Number(params.indice);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_BALANCE]) === String(params.idBalance)) {
      const observaciones = JSON.parse(datos[i][idx.OBSERVACIONES]);
      if (!observaciones[indice]) return { ok: false, error: 'Observación no encontrada' };

      observaciones[indice].estadoObs = params.estadoObs;
      observaciones[indice].comentario = params.comentario || observaciones[indice].comentario || '';
      if (params.estadoObs === 'PENDIENTE') {
        observaciones[indice].marcadaPor = '';
        observaciones[indice].marcadaFecha = '';
      } else {
        observaciones[indice].marcadaPor = params.pin || '';
        observaciones[indice].marcadaFecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      }

      hoja.getRange(i + 1, idx.OBSERVACIONES + 1).setValue(JSON.stringify(observaciones));

      // Refleja el cambio también dentro de DATOS_EXTRAIDOS, para que quede consistente
      const datosExtraidos = JSON.parse(datos[i][idx.DATOS_EXTRAIDOS]);
      if (datosExtraidos.observaciones && datosExtraidos.observaciones[indice]) {
        datosExtraidos.observaciones[indice].estadoObs = params.estadoObs;
        datosExtraidos.observaciones[indice].comentario = observaciones[indice].comentario;
        hoja.getRange(i + 1, idx.DATOS_EXTRAIDOS + 1).setValue(JSON.stringify(datosExtraidos));
      }

      // Si ya no quedan CRITICOS pendientes, sugiere avanzar de OBSERVADO a EN_REVISION automáticamente
      const quedanCriticosPendientes = observaciones.some(o => o.tipo === 'CRITICO' && o.estadoObs === 'PENDIENTE');
      let nuevoEstadoGeneral = datos[i][idx.ESTADO];
      if (!quedanCriticosPendientes && nuevoEstadoGeneral === 'OBSERVADO') {
        nuevoEstadoGeneral = 'EN_REVISION';
        hoja.getRange(i + 1, idx.ESTADO + 1).setValue(nuevoEstadoGeneral);
      }

      return { ok: true, observaciones: observaciones, estadoGeneral: nuevoEstadoGeneral };
    }
  }
  return { ok: false, error: 'Balance no encontrado' };
}

// ============ EJERCICIOS ============
const HOJA_EJERCICIOS = 'EJERCICIOS';

function getHojaEjercicios() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_EJERCICIOS);
}

function listarEjercicios() {
  const hoja = getHojaEjercicios();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const ejercicios = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_EJERCICIO]) continue;
    ejercicios.push({
      idEjercicio: fila[idx.ID_EJERCICIO],
      numero: fila[idx.NUMERO],
      fechaInicio: Utilities.formatDate(new Date(fila[idx.FECHA_INICIO]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      fechaCierre: Utilities.formatDate(new Date(fila[idx.FECHA_CIERRE]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      estado: fila[idx.ESTADO]
    });
  }
  ejercicios.sort((a, b) => Number(b.numero) - Number(a.numero));
  return { ok: true, ejercicios: ejercicios };
}

// Resumen agregado para mostrar en la tile del Centro de Control -- una sola llamada, no cinco
function obtenerResumenDashboard() {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'Sin Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  // Autoridades: vacantes sobre 12 cargos de Comisión Directiva
  const autRes = listarAutoridades();
  const vacantesAutoridades = autRes.ok ? Math.max(0, 12 - autRes.autoridades.length) : null;
  const vencenEsteAnio = autRes.ok ? autRes.vencenEsteAnio.length : 0;

  // Actas: cuántas en borrador
  const actasRes = listarActas();
  const actasBorrador = actasRes.ok ? actasRes.actas.filter(a => a.estado === 'BORRADOR').length : 0;
  const actasTotal = actasRes.ok ? actasRes.actas.length : 0;

  // Balance: estado del último del ejercicio activo
  const balancesRes = listarBalances();
  const balanceDelEjercicio = balancesRes.ok ? balancesRes.balances.find(b => b.idEjercicio === ej.idEjercicio) : null;

  // Novedades: cuántas sin revisar (EVALUAR)
  const novedadesRes = listarNovedades({ idEjercicio: ej.idEjercicio });
  const novedadesPendientes = novedadesRes.ok ? novedadesRes.novedades.filter(n => n.considerarMemoria === 'EVALUAR').length : 0;
  const novedadesAprobadas = novedadesRes.ok ? novedadesRes.novedades.filter(n => n.considerarMemoria === 'SI').length : 0;

  // Memoria: última versión generada
  const memoriaRes = listarDocumentos({ tipo: 'MEMORIA', idEjercicio: ej.idEjercicio });
  const ultimaMemoria = memoriaRes.ok && memoriaRes.documentos.length ? memoriaRes.documentos[0] : null;

  // Progreso: mismos 10 hitos y pesos que usa el Dashboard de la app (3 calculados en vivo, 7 con el estado real conocido hoy)
  const pesoEstado = { COMPLETO: 1, PROGRESO: 0.5, OBSERVADO: 0.25, PENDIENTE: 0, NOINICIADO: 0 };
  const estadoBalance = !balanceDelEjercicio ? 'NOINICIADO' : (balanceDelEjercicio.estado === 'APROBADO_ASAMBLEA' ? 'COMPLETO' : (balanceDelEjercicio.estado === 'OBSERVADO' ? 'OBSERVADO' : 'PROGRESO'));
  const estadoAutoridades = vacantesAutoridades === 0 ? 'COMPLETO' : 'PROGRESO';
  const estadoActas = actasTotal === 0 ? 'NOINICIADO' : (actasBorrador > 0 ? 'PROGRESO' : 'COMPLETO');
  const hitosFijos = ['COMPLETO', 'PENDIENTE', ultimaMemoria ? 'PROGRESO' : 'PENDIENTE', 'PROGRESO', 'PROGRESO', 'NOINICIADO', 'PROGRESO'];
  const todosLosPesos = [estadoBalance, estadoAutoridades, estadoActas].concat(hitosFijos).map(e => pesoEstado[e]);
  const progresoPct = Math.round((todosLosPesos.reduce((a, b) => a + b, 0) / todosLosPesos.length) * 100);

  return {
    ok: true,
    ejercicioNumero: ej.numero,
    ejercicioFechas: ej.fechaInicio + ' a ' + ej.fechaCierre,
    progresoPct: progresoPct,
    balanceEstado: balanceDelEjercicio ? balanceDelEjercicio.estado : 'SIN_PROCESAR',
    autoridadesVacantes: vacantesAutoridades,
    autoridadesVencenEsteAnio: vencenEsteAnio,
    actasBorrador: actasBorrador,
    actasTotal: actasTotal,
    novedadesPendientes: novedadesPendientes,
    novedadesAprobadas: novedadesAprobadas,
    memoriaVersion: ultimaMemoria ? ultimaMemoria.version : 0,
    memoriaEstado: ultimaMemoria ? ultimaMemoria.estado : 'SIN_GENERAR'
  };
}

function obtenerEjercicioActivo() {
  const hoja = getHojaEjercicios();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  // El activo es el de mayor número que no esté PRESENTADO
  let activo = null;
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_EJERCICIO]) continue;
    if (fila[idx.ESTADO] !== 'PRESENTADO') {
      if (!activo || Number(fila[idx.NUMERO]) > Number(activo[idx.NUMERO])) activo = fila;
    }
  }
  if (!activo) return { ok: false, error: 'No hay ningún Ejercicio activo cargado' };

  return {
    ok: true,
    ejercicio: {
      idEjercicio: activo[idx.ID_EJERCICIO],
      numero: activo[idx.NUMERO],
      fechaInicio: Utilities.formatDate(new Date(activo[idx.FECHA_INICIO]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      fechaCierre: Utilities.formatDate(new Date(activo[idx.FECHA_CIERRE]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      fechaLimiteAsamblea: Utilities.formatDate(new Date(activo[idx.FECHA_LIMITE_ASAMBLEA]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      estado: activo[idx.ESTADO]
    }
  };
}

function cerrarYAbrirNuevoEjercicio(params) {
  const hoja = getHojaEjercicios();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let filaActiva = -1;
  let activo = null;
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_EJERCICIO]) continue;
    if (fila[idx.ESTADO] !== 'PRESENTADO') {
      if (!activo || Number(fila[idx.NUMERO]) > Number(activo[idx.NUMERO])) { activo = fila; filaActiva = i; }
    }
  }
  if (!activo) return { ok: false, error: 'No hay Ejercicio activo para cerrar' };

  // Cierra el actual
  hoja.getRange(filaActiva + 1, idx.ESTADO + 1).setValue('PRESENTADO');

  // Abre el siguiente: inicio = día después del cierre anterior, cierre = +1 año, límite asamblea = cierre + 3 meses
  const cierreAnterior = new Date(activo[idx.FECHA_CIERRE]);
  const nuevoInicio = new Date(cierreAnterior);
  nuevoInicio.setDate(nuevoInicio.getDate() + 1);
  const nuevoCierre = new Date(nuevoInicio);
  nuevoCierre.setFullYear(nuevoCierre.getFullYear() + 1);
  nuevoCierre.setDate(nuevoCierre.getDate() - 1);
  const nuevoLimiteAsamblea = new Date(nuevoCierre);
  nuevoLimiteAsamblea.setMonth(nuevoLimiteAsamblea.getMonth() + 3);

  const nuevoNumero = Number(activo[idx.NUMERO]) + 1;
  const nuevoId = 'EJ-' + String(nuevoNumero).padStart(3, '0');

  hoja.appendRow([nuevoId, nuevoNumero, nuevoInicio, nuevoCierre, 'ABIERTO', nuevoLimiteAsamblea]);

  return { ok: true, ejercicioAnteriorCerrado: activo[idx.ID_EJERCICIO], nuevoEjercicio: nuevoId, nuevoNumero: nuevoNumero };
}

// ============ NOVEDADES ============
const HOJA_NOVEDADES = 'NOVEDADES';

// Regla por defecto: Personal nunca se marca SI automáticamente (ver criterio de Memoria)
const CONSIDERAR_MEMORIA_DEFAULT = { 'Personal': 'NO' };

// Nunca deja guardar una novedad con fecha fuera del período del Ejercicio activo.
// El pasado ya cerrado no se toca desde acá; el futuro (próximo Ejercicio) todavía no existe como fila.
// Convierte 'yyyy-MM-dd' a Date sin el corrimiento de un día que da 'new Date(string)' por interpretarlo como UTC
function parsearFechaSegura(fechaStr) {
  const partes = fechaStr.split('-');
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
}

function fechaEnRangoEjercicio(fechaStr, ejercicio) {
  const fecha = parsearFechaSegura(fechaStr.split('/').reverse().join('-')); // fechaStr viene dd/mm/aaaa
  const inicio = parsearFechaSegura(ejercicio.fechaInicio.split('/').reverse().join('-'));
  const cierre = parsearFechaSegura(ejercicio.fechaCierre.split('/').reverse().join('-'));
  if (fecha < inicio || fecha > cierre) {
    return {
      ok: false,
      error: 'La fecha ' + fechaStr + ' está fuera del período del Ejercicio N.° ' + ejercicio.numero +
        ' (' + ejercicio.fechaInicio + ' a ' + ejercicio.fechaCierre + '). No se guarda para evitar mezclar ejercicios.'
    };
  }
  return { ok: true };
}

function getHojaNovedades() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_NOVEDADES);
}

function guardarNovedad(params) {
  const hoja = getHojaNovedades();
  const nuevoId = 'NOV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

  const ejercicioActivo = obtenerEjercicioActivo();
  const idEjercicioActual = ejercicioActivo.ok ? ejercicioActivo.ejercicio.idEjercicio : '';

  if (ejercicioActivo.ok) {
    const chequeo = fechaEnRangoEjercicio(params.fecha, ejercicioActivo.ejercicio);
    if (!chequeo.ok) return chequeo;
  }

  const considerarMemoria = params.considerarMemoria || CONSIDERAR_MEMORIA_DEFAULT[params.categoria] || 'EVALUAR';

  hoja.appendRow([
    nuevoId,
    idEjercicioActual,
    parsearFechaSegura(params.fecha),
    params.titulo,
    params.descripcion || '',
    params.categoria,
    params.monto || '',
    'MANUAL',
    '',
    considerarMemoria,
    params.responsable || ''
  ]);

  return { ok: true, idNovedad: nuevoId };
}

// ============ SINCRONIZACIÓN CON EL BOLETÍN MENSUAL ============
// La cola de "Novedades del mes" del Boletín es una fuente real de hechos institucionales.
// Esto la lee, descarta lo que ya se importó antes (por ORIGEN=BOLETIN + ID_ORIGEN), y devuelve solo lo nuevo para revisar.
function listarNovedadesDesdeBoletin() {
  const ssBoletin = SpreadsheetApp.openById(SHEET_ID_BOLETIN);
  const hojas = ssBoletin.getSheets();
  let hojaCola = null;
  let idx = {};
  for (const hoja of hojas) {
    const datos = hoja.getDataRange().getValues();
    if (datos.length === 0) continue;
    const headers = datos[0].map(h => String(h).trim().toUpperCase());
    if (headers.includes('ID') && headers.includes('TITULO') && headers.includes('DESCRIPCION') && headers.includes('FECHA_CARGA')) {
      hojaCola = hoja;
      headers.forEach((h, i) => idx[h] = i);
      break;
    }
  }
  if (!hojaCola) return { ok: false, error: 'No se encontró la hoja de Novedades del Boletín (buscando columnas ID/TITULO/DESCRIPCION/FECHA_CARGA)' };

  const datosCola = hojaCola.getDataRange().getValues();
  const candidatos = [];
  for (let i = 1; i < datosCola.length; i++) {
    const fila = datosCola[i];
    if (!fila[idx.ID]) continue;
    candidatos.push({
      idBoletin: fila[idx.ID],
      titulo: fila[idx.TITULO],
      descripcion: fila[idx.DESCRIPCION],
      fechaCarga: fila[idx.FECHA_CARGA] instanceof Date ? Utilities.formatDate(fila[idx.FECHA_CARGA], Session.getScriptTimeZone(), 'dd/MM/yyyy') : fila[idx.FECHA_CARGA]
    });
  }

  // Filtra lo que ya se importó antes
  const hojaNov = getHojaNovedades();
  const datosNov = hojaNov.getDataRange().getValues();
  const headersNov = datosNov[0];
  const idxNov = {};
  headersNov.forEach((h, i) => idxNov[h] = i);
  const yaImportados = new Set();
  for (let i = 1; i < datosNov.length; i++) {
    if (datosNov[i][idxNov.ORIGEN] === 'BOLETIN') yaImportados.add(String(datosNov[i][idxNov.ID_ORIGEN]));
  }

  const nuevos = candidatos.filter(c => !yaImportados.has(String(c.idBoletin)));
  return { ok: true, novedades: nuevos, totalEnBoletin: candidatos.length };
}

function listarNovedades(params) {
  const hoja = getHojaNovedades();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const filtroEjercicio = params && params.idEjercicio ? params.idEjercicio : null;

  const novedades = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_NOVEDAD]) continue;
    if (filtroEjercicio && fila[idx.ID_EJERCICIO] !== filtroEjercicio) continue;
    novedades.push({
      idNovedad: fila[idx.ID_NOVEDAD],
      idEjercicio: fila[idx.ID_EJERCICIO],
      fecha: fila[idx.FECHA] ? Utilities.formatDate(new Date(fila[idx.FECHA]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      fechaISO: fila[idx.FECHA] ? Utilities.formatDate(new Date(fila[idx.FECHA]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      titulo: fila[idx.TITULO],
      descripcion: fila[idx.DESCRIPCION],
      categoria: fila[idx.CATEGORIA],
      monto: fila[idx.MONTO],
      origen: fila[idx.ORIGEN],
      considerarMemoria: fila[idx.CONSIDERAR_MEMORIA],
      responsable: fila[idx.RESPONSABLE]
    });
  }
  novedades.sort((a, b) => new Date(b.fecha.split('/').reverse().join('-')) - new Date(a.fecha.split('/').reverse().join('-')));
  return { ok: true, novedades: novedades };
}

function actualizarNovedad(params) {
  const hoja = getHojaNovedades();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_NOVEDAD]) === String(params.idNovedad)) {
      // Si se cambia la fecha, valida que siga cayendo dentro del Ejercicio al que ya pertenece esta novedad
      if (params.fecha) {
        const idEjercicioNovedad = datos[i][idx.ID_EJERCICIO];
        const ejercicios = listarEjercicios();
        const ejercicioDeLaNovedad = ejercicios.ok ? ejercicios.ejercicios.find(e => e.idEjercicio === idEjercicioNovedad) : null;
        if (ejercicioDeLaNovedad) {
          const chequeo = fechaEnRangoEjercicio(params.fecha, ejercicioDeLaNovedad);
          if (!chequeo.ok) return chequeo;
        }
        hoja.getRange(i + 1, idx.FECHA + 1).setValue(parsearFechaSegura(params.fecha));
      }
      if (params.considerarMemoria) hoja.getRange(i + 1, idx.CONSIDERAR_MEMORIA + 1).setValue(params.considerarMemoria);
      if (params.titulo) hoja.getRange(i + 1, idx.TITULO + 1).setValue(params.titulo);
      if (params.descripcion) hoja.getRange(i + 1, idx.DESCRIPCION + 1).setValue(params.descripcion);
      if (params.categoria) hoja.getRange(i + 1, idx.CATEGORIA + 1).setValue(params.categoria);
      if (params.monto !== undefined) hoja.getRange(i + 1, idx.MONTO + 1).setValue(params.monto);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Novedad no encontrada' };
}

function eliminarNovedad(params) {
  const hoja = getHojaNovedades();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_NOVEDAD]) === String(params.idNovedad)) {
      hoja.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Novedad no encontrada' };
}

const PROMPT_EXTRAER_NOVEDADES = `Sos un asistente que ayuda a una asociación civil (Club de Campo "La Eugenia") a separar información institucional real de charla social/ruido dentro de exportaciones de chat de WhatsApp (grupo de Comisión Directiva o grupo de difusión a socios).

Vas a recibir texto crudo de un chat exportado (con fechas, remitentes, mensajes, emojis, stickers, "<Multimedia omitido>", etc.).

Tu tarea: identificar HECHOS o DECISIONES institucionales concretos que valga la pena registrar como "Novedad" del Club -- obras, gestiones ante organismos (Municipalidad, EMSA, SAMSA, Personas Jurídicas), eventos, gastos relevantes, incidentes de seguridad, decisiones de la Comisión Directiva, cambios de servicios (cantina, pileta, etc.). IGNORÁ: saludos, charla social, memes, stickers, mensajes ambiguos sin sustancia institucional, y cualquier chisme o comentario personal sobre empleados o socios.

Para cada hecho identificado, respondé un objeto con:
{
  "fecha": "DD/MM/AAAA (la fecha real del mensaje en el chat, no la de hoy)",
  "titulo": "resumen corto, 5-8 palabras",
  "descripcion": "1-3 oraciones, en tono institucional, redactado por vos en base a lo que dice el chat -- no copies el mensaje textual, resumilo",
  "categoria": "una de: Administración, Seguridad, Mantenimiento, Infraestructura, Alumbrado, Pileta, Deportes, Actividades sociales, Cantina, Personal, Municipalidad, EMSA, SAMSA, Personas Jurídicas, Tesorería, Incidentes, Otros",
  "montoSiCorresponde": "número sin signos, o vacío si no aplica"
}

Respondé ÚNICAMENTE con un array JSON de estos objetos, sin texto adicional, sin markdown. Si no hay nada relevante, respondé un array vacío [].`;

function extraerNovedadesDeChat(params) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'Falta configurar ANTHROPIC_API_KEY en Script Properties' };

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 8000,
    thinking: { type: 'disabled' },
    system: PROMPT_EXTRAER_NOVEDADES,
    messages: [{ role: 'user', content: params.textoChat }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return { ok: false, error: 'Error de la API (' + response.getResponseCode() + '): ' + response.getContentText() };
  }

  const data = JSON.parse(response.getContentText());
  const bloqueTexto = (data.content || []).find(b => b.type === 'text');
  if (!bloqueTexto) return { ok: false, error: 'La IA no devolvió texto' };

  let texto = bloqueTexto.text.replace(/\`\`\`json|\`\`\`/g, '').trim();
  let candidatos;
  try {
    candidatos = JSON.parse(texto);
  } catch (e) {
    return { ok: false, error: 'La IA no devolvió JSON válido: ' + texto.substring(0, 300) };
  }

  // Marca posibles duplicados comparando contra lo ya cargado (mismo período aproximado + título parecido)
  const ejercicioActivo = obtenerEjercicioActivo();
  const existentes = ejercicioActivo.ok ? (listarNovedades({ idEjercicio: ejercicioActivo.ejercicio.idEjercicio }).novedades || []) : [];

  candidatos.forEach(c => {
    const posible = existentes.find(e => sonProbablesDuplicados(c, e));
    if (posible) {
      c.posibleDuplicadoDe = posible.titulo + ' (' + posible.fecha + ')';
    }
  });

  return { ok: true, candidatos: candidatos };
}

function normalizarTexto(s) {
  return (s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // saca acentos
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/).filter(w => w.length > 3); // palabras significativas
}

function sonProbablesDuplicados(candidato, existente) {
  // Fecha cercana (7 días) + al menos 3 palabras significativas del título en común
  const fechaCand = parsearFechaSegura(candidato.fecha.split('/').reverse().join('-'));
  const fechaExist = parsearFechaSegura(existente.fecha.split('/').reverse().join('-'));
  const diffDias = Math.abs((fechaCand - fechaExist) / (1000 * 60 * 60 * 24));
  if (diffDias > 7) return false;

  const palabrasCand = new Set(normalizarTexto(candidato.titulo));
  const palabrasExist = new Set(normalizarTexto(existente.titulo));
  let comunes = 0;
  palabrasCand.forEach(w => { if (palabrasExist.has(w)) comunes++; });

  return comunes >= 3;
}

function guardarNovedadesSeleccionadas(params) {
  const seleccionadas = JSON.parse(params.novedades);
  const hoja = getHojaNovedades();
  const ejercicioActivo = obtenerEjercicioActivo();
  const idEjercicioActual = ejercicioActivo.ok ? ejercicioActivo.ejercicio.idEjercicio : '';

  let guardadas = 0;
  const rechazadas = [];
  seleccionadas.forEach(n => {
    if (ejercicioActivo.ok) {
      const chequeo = fechaEnRangoEjercicio(n.fecha, ejercicioActivo.ejercicio);
      if (!chequeo.ok) { rechazadas.push({ titulo: n.titulo, motivo: chequeo.error }); return; }
    }
    const nuevoId = 'NOV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + guardadas;
    const considerarMemoria = CONSIDERAR_MEMORIA_DEFAULT[n.categoria] || 'EVALUAR';
    hoja.appendRow([
      nuevoId, idEjercicioActual, parsearFechaSegura(n.fecha.split('/').reverse().join('-')),
      n.titulo, n.descripcion, n.categoria, n.montoSiCorresponde || '',
      n.origen || 'CHAT_IMPORTADO', n.idOrigen || '', considerarMemoria, ''
    ]);
    guardadas++;
  });

  return { ok: true, guardadas: guardadas, rechazadas: rechazadas };
}

// ============ DOCUMENTOS (tabla única: Memoria, y a futuro Convocatoria/Edictos/Informe Revisor, etc.) ============
const HOJA_DOCUMENTOS = 'DOCUMENTOS';

function getHojaDocumentos() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_DOCUMENTOS);
}

const PROMPT_MEMORIA = `Sos el redactor institucional del Club de Campo "La Eugenia", una asociación civil sin fines de lucro. Vas a escribir el borrador de la MEMORIA anual, el documento que el Presidente lee y la Comisión Directiva presenta a la Asamblea General Ordinaria, en cumplimiento del Estatuto Social.

ESTILO OBLIGATORIO (aprendido de memorias reales anteriores del Club, no te apartes de esto):
- Arranca siempre EXACTAMENTE así: "Señores Socios:\\n\\nEn cumplimiento de las disposiciones legales y estatutarias, ponemos a vuestra consideración la presente memoria, junto con los estados contables, el informe del auditor y el informe del revisor de cuentas correspondientes al Ejercicio Económico N.° [NUMERO DE EJERCICIO], finalizado el [FECHA DE CIERRE]." IMPORTANTE: siempre incluí el número de Ejercicio explícitamente en esta primera frase (ej. "Ejercicio Económico N.° 37") -- es necesario para diferenciar esta Memoria de las de otros años, aunque las memorias anteriores no siempre lo hayan incluido.
- El cuerpo es PROSA CORRIDA, en párrafos, SIN subtítulos ni encabezados de sección. Nunca uses títulos tipo "Seguridad:" o "Mantenimiento:" — los temas se mezclan en párrafos fluidos, ordenados por relevancia.
- Tono: primera persona plural ("hemos", "continuamos", "nos comprometemos"), formal pero cálido, nunca frío ni corporativo.
- Los temas se agrupan naturalmente en el texto (infraestructura, seguridad, EMSA/alumbrado, mantenimiento de espacios verdes, temporada de pileta, cantina, actividades sociales/deportivas, gestiones ante organismos como Municipalidad/EMSA/SAMSA/Personas Jurídicas), pero SIN anunciarlos como secciones.
- Los temas de "Personal" (altas, bajas, retribuciones, conflictos laborales) NUNCA se mencionan en la Memoria, aunque estén en las Novedades — no son parte de este documento.
- Cierra siempre con el desglose de socios que te paso como dato, con esta fórmula: "A la fecha de la presente memoria el número de socios Activos asciende a [N], de los cuales [N] abonan al día, mientras que [N] abona con un atraso de entre uno y cinco meses. Los socios No Activos, Morosos, son [N]."
- Después, SIEMPRE el agradecimiento: "La Comisión Directiva agradece profundamente a todos quienes colaboran con la gestión institucional y con el crecimiento sostenido del club, en especial al personal, cuyo compromiso y dedicación resultan fundamentales para el logro de los objetivos."
- Cierra con: "Garupá, [FECHA].\\n\\nCOMISIÓN DIRECTIVA"

Te paso las Novedades del ejercicio (ya filtradas a las que corresponde incluir) y los datos del Balance. Redactá el cuerpo integrando la información de forma natural y fluida, sin inventar nada que no esté en los datos que te doy. Si falta información importante para algún párrafo, simplemente no lo incluyas -- no inventes cifras ni hechos.

Respondé ÚNICAMENTE con el texto completo de la Memoria, sin explicaciones adicionales, sin markdown.`;

// ============ CONFIG MEMORIA (editable por el usuario, sin tocar código) ============
const HOJA_CONFIG_MEMORIA = 'MEMORIA_CONFIG';

function getHojaConfigMemoria() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hoja = ss.getSheetByName(HOJA_CONFIG_MEMORIA);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_CONFIG_MEMORIA);
  }
  return hoja;
}

function obtenerConfigMemoria() {
  const valor = obtenerValorPorClave('INSTRUCCIONES_ADICIONALES');
  return { ok: true, instrucciones: valor || '' };
}

function guardarConfigMemoria(params) {
  guardarValorPorClave('INSTRUCCIONES_ADICIONALES', params.instrucciones || '');
  return { ok: true };
}

// Clave/valor genérico en MEMORIA_CONFIG -- se reutiliza para instrucciones de estilo,
// últimos números de socios cargados por Ejercicio, y lo que haga falta a futuro.
function obtenerValorPorClave(clave) {
  const hoja = getHojaConfigMemoria();
  const datos = hoja.getDataRange().getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === clave) return datos[i][1];
  }
  return null;
}

function guardarValorPorClave(clave, valor) {
  const hoja = getHojaConfigMemoria();
  const datos = hoja.getDataRange().getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === clave) {
      hoja.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  hoja.appendRow([clave, valor]);
}

function obtenerUltimosSocios(params) {
  const valor = obtenerValorPorClave('SOCIOS_' + params.idEjercicio);
  if (!valor) return { ok: true, socios: null };
  try {
    return { ok: true, socios: JSON.parse(valor) };
  } catch (e) {
    return { ok: true, socios: null };
  }
}

const CARPETA_INFORMES_REVISOR_ID = '1hI8nzHEraUxqEoKAZdqjjtdJKj9FTWnO'; // Informes_Revisor_Gestion_Institucional_La_Eugenia

const PROMPT_RESUMEN_INFORME_REVISOR = `Vas a recibir el Informe de la Comisión Revisora de Cuentas de una asociación civil, ya redactado y firmado por el Revisor. NO lo corrijas ni lo cuestiones -- el Revisor es la autoridad sobre este documento, no vos. Tu única tarea es extraer un resumen breve para que quede registrado en el sistema.

Respondé ÚNICAMENTE con un JSON:
{
  "fechaInforme": "DD/MM/AAAA si figura en el documento, o vacío",
  "conclusion": "APRUEBA_SIN_SALVEDADES o APRUEBA_CON_SALVEDADES o NO_APRUEBA o INDETERMINADO",
  "resumen": "1-2 oraciones resumiendo qué dice el informe, en tono neutral"
}`;

// ============ CONVOCATORIA Y DOCUMENTOS DE ASAMBLEA ============
const HOJA_ASAMBLEAS = 'ASAMBLEAS';

function getHojaAsambleas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hoja = ss.getSheetByName(HOJA_ASAMBLEAS);
  if (!hoja) hoja = ss.insertSheet(HOJA_ASAMBLEAS);
  return hoja;
}

// La app SIEMPRE debe leer la fecha de acá, nunca tenerla escrita fija en el código del front-end.
// Si Sheets convirtió el texto "16:00" en un objeto de hora, lo devuelve al formato HH:mm; si ya es texto, lo deja igual
function formatearHoraSegura(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'HH:mm');
  }
  return valor;
}

function obtenerAsambleaEjercicio() {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };

  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) {
      let ordenDelDia = [];
      try { ordenDelDia = JSON.parse(datos[i][5]); } catch (e) {}
      return {
        ok: true,
        idAsamblea: datos[i][0],
        fecha: Utilities.formatDate(new Date(datos[i][2]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        hora: formatearHoraSegura(datos[i][3]),
        lugar: datos[i][4],
        estado: datos[i][6],
        ordenDelDia: ordenDelDia,
        autoridadesRegistradas: datos[i][7] === 'SI'
      };
    }
  }
  return { ok: false, error: 'Todavía no hay fecha de Asamblea registrada para este Ejercicio' };
}

// Guarda o actualiza fecha/hora/lugar de la Asamblea del Ejercicio activo -- editable desde el Dashboard, dato de primera clase
// Corrige puntualmente la fecha ya guardada con el bug de huso horario (quedó en 23/10 en vez de 24/10)
function diagnosticoAsambleas() {
  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  const ejercicioActivo = obtenerEjercicioActivo();
  return {
    ok: true,
    idEjercicioActivo: ejercicioActivo.ok ? ejercicioActivo.ejercicio.idEjercicio : 'SIN EJERCICIO ACTIVO',
    filas: datos.map(fila => fila.map(celda => JSON.stringify(celda)))
  };
}

// Limpia duplicados en ASAMBLEAS (bug ya corregido, esto es solo para arreglar lo que ya quedó mal)
function limpiarDuplicadosAsambleas() {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();

  const filasDelEjercicio = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) filasDelEjercicio.push(i + 1);
  }
  if (filasDelEjercicio.length <= 1) return { ok: true, mensaje: 'No había duplicados' };

  // Deja la última (la más reciente) y borra las anteriores
  const filaAConservar = filasDelEjercicio[filasDelEjercicio.length - 1];
  const aBorrar = filasDelEjercicio.slice(0, -1).sort((a, b) => b - a); // de mayor a menor, para no correr los índices al borrar
  aBorrar.forEach(f => hoja.deleteRow(f));

  return { ok: true, borradas: aBorrar.length, conservada: filaAConservar };
}

function corregirFechaAsambleaExistente() {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() === ejercicioActivo.ejercicio.idEjercicio) {
      hoja.getRange(i + 1, 3).setValue(parsearFechaSegura('2026-10-24'));
      return { ok: true, filaEncontrada: i + 1 };
    }
  }
  return { ok: false, error: 'No se encontró ninguna fila con ID_EJERCICIO=' + ejercicioActivo.ejercicio.idEjercicio + ' en ASAMBLEAS' };
}

function guardarAsambleaEjercicio(params) {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };

  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) {
      hoja.getRange(i + 1, 3).setValue(parsearFechaSegura(params.fecha));
      const celdaHora = hoja.getRange(i + 1, 4);
      celdaHora.setNumberFormat('@'); // fuerza texto plano, para que Sheets no la convierta en hora
      celdaHora.setValue(params.hora);
      hoja.getRange(i + 1, 5).setValue(params.lugar || '');
      return { ok: true };
    }
  }
  const nuevaFila = hoja.getLastRow() + 1;
  hoja.appendRow(['ASM-' + ejercicioActivo.ejercicio.numero, ejercicioActivo.ejercicio.idEjercicio, parsearFechaSegura(params.fecha), params.hora, params.lugar || '', JSON.stringify([]), 'CONVOCADA', '', '', '']);
  hoja.getRange(nuevaFila, 4).setNumberFormat('@').setValue(params.hora); // refuerza formato de texto en la fila nueva
  return { ok: true };
}

// Carga puntual del dato real ya decidido en el Acta 562 -- una sola vez, para no dejar la hoja vacía
function sembrarAsambleaReal() {
  const ejercicioActivo = obtenerEjercicioActivo();
  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) return { ok: true, mensaje: 'Ya existía, no se duplica' };
  }
  hoja.appendRow(['ASM-' + ejercicioActivo.ejercicio.numero, ejercicioActivo.ejercicio.idEjercicio, new Date('2026-10-24'), '16:00', 'S.U.M. del Club, Ruta Nacional N.° 105, Km 5, Garupá, Misiones', JSON.stringify([]), 'CONVOCADA', '', '', '']);
  return { ok: true };
}

function calcularFechaLimiteAsamblea(fechaCierreStr) {
  const cierre = new Date(fechaCierreStr.split('/').reverse().join('-'));
  const limite = new Date(cierre);
  limite.setMonth(limite.getMonth() + 3);
  return limite;
}

// ============ ASAMBLEA (el día del evento) ============

function marcarAsambleaCelebrada(params) {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) {
      hoja.getRange(i + 1, 7).setValue('CELEBRADA');
      hoja.getRange(i + 1, 8).setValue(params.sociosPresentes || '');
      hoja.getRange(i + 1, 9).setValue(params.quorumAlcanzado || '');
      return { ok: true };
    }
  }
  return { ok: false, error: 'No hay Asamblea registrada para este Ejercicio' };
}

function registrarAutoridadesElectas(params) {
  const elecciones = JSON.parse(params.elecciones); // [{cargo, grupo, nombreNuevo}]
  const hoja = getHoja();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const hoy = new Date();
  const finMandato = new Date(hoy);
  finMandato.setFullYear(finMandato.getFullYear() + 2);

  let actualizadas = 0;
  elecciones.forEach(e => {
    // Finaliza al titular anterior de ese cargo, si está VIGENTE
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][idx.CARGO] === e.cargo && datos[i][idx.ESTADO] === 'VIGENTE') {
        hoja.getRange(i + 1, idx.ESTADO + 1).setValue('FINALIZADO');
      }
    }
    // Da de alta al nuevo, si se informó un nombre
    if (e.nombreNuevo && e.nombreNuevo.trim()) {
      hoja.appendRow(['AUT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + actualizadas, 'COMISION_DIRECTIVA', e.cargo, e.grupo || '', e.nombreNuevo, '', hoy, finMandato, 'VIGENTE', '']);
      actualizadas++;
    }
  });

  // Marca en ASAMBLEAS que ya se registraron autoridades, para el cálculo de progreso del Dashboard
  const ejercicioActivo = obtenerEjercicioActivo();
  if (ejercicioActivo.ok) {
    const hojaAsam = getHojaAsambleas();
    const datosAsam = hojaAsam.getDataRange().getValues();
    for (let i = 1; i < datosAsam.length; i++) {
      if (datosAsam[i][1] === ejercicioActivo.ejercicio.idEjercicio) {
        hojaAsam.getRange(i + 1, 8).setValue('SI'); // columna 8 = AUTORIDADES_REGISTRADAS
        break;
      }
    }
  }

  return { ok: true, actualizadas: actualizadas };
}

const PROMPT_ACTA_ASAMBLEA = `Sos el redactor institucional del Club de Campo "La Eugenia". Vas a escribir el ACTA DE ASAMBLEA GENERAL ORDINARIA, el documento formal que registra lo tratado y resuelto en la Asamblea.

ESTILO OBLIGATORIO:
- Arranca así: "ACTA DE ASAMBLEA GENERAL ORDINARIA\\n\\nEn la ciudad de Garupá, Provincia de Misiones, siendo las [HORA] horas del día [FECHA], se reúnen los socios del Club de Campo \"La Eugenia\" en Asamblea General Ordinaria, en [LUGAR], para tratar el Orden del Día previamente notificado."
- Menciona el quórum alcanzado y los socios presentes según el dato que te paso.
- Desarrollá cada punto del Orden del Día como un párrafo de resultado (aprobado, con modificaciones, etc.), usando los datos reales que te doy -- no inventes resultados que no te haya dado.
- Si se informaron autoridades electas, mencionalas nominalmente en el punto de renovación de Comisión Directiva.
- Cierra con: "Siendo las [HORA] horas se da por finalizada la Asamblea, firmando la presente el Presidente y el Secretario de Asamblea designados.\\n\\nGarupá, [FECHA]."

Respondé ÚNICAMENTE con el texto completo del acta, sin explicaciones adicionales, sin markdown.`;

function generarActaAsamblea(params) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'Falta configurar ANTHROPIC_API_KEY en Script Properties' };

  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const asamblea = obtenerAsambleaEjercicio();
  if (!asamblea.ok) return { ok: false, error: 'No hay Asamblea registrada' };

  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  let ordenDelDia = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) {
      try { ordenDelDia = JSON.parse(datos[i][5]); } catch (e) {}
      break;
    }
  }

  const inputTexto = 'FECHA: ' + asamblea.fecha + '\nHORA: ' + asamblea.hora + '\nLUGAR: ' + asamblea.lugar +
    '\nSOCIOS PRESENTES: ' + (params.sociosPresentes || '[no informado]') + '\nQUÓRUM: ' + (params.quorumAlcanzado || '[no informado]') +
    '\n\nORDEN DEL DÍA:\n' + ordenDelDia.join('\n') +
    '\n\nRESULTADOS INFORMADOS POR PUNTO:\n' + (params.resultados || '[no informado]') +
    '\n\nAUTORIDADES ELECTAS (si corresponde):\n' + (params.autoridadesElectas || '[no informado]');

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 6000,
    thinking: { type: 'disabled' },
    system: PROMPT_ACTA_ASAMBLEA,
    messages: [{ role: 'user', content: inputTexto }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return { ok: false, error: 'Error de la API (' + response.getResponseCode() + '): ' + response.getContentText() };
  }
  const data = JSON.parse(response.getContentText());
  const bloqueTexto = (data.content || []).find(b => b.type === 'text');
  if (!bloqueTexto) return { ok: false, error: 'La IA no devolvió texto' };

  const hojaDocs = getHojaDocumentos();
  const datosDocs = hojaDocs.getDataRange().getValues();
  let version = 1;
  for (let i = 1; i < datosDocs.length; i++) {
    if (datosDocs[i][0] && datosDocs[i][2] === 'ACTA_ASAMBLEA' && datosDocs[i][1] === ejercicioActivo.ejercicio.idEjercicio) version++;
  }
  const nuevoId = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  hojaDocs.appendRow([nuevoId, ejercicioActivo.ejercicio.idEjercicio, 'ACTA_ASAMBLEA', version, 'BORRADOR', bloqueTexto.text, 'IA', new Date()]);

  return { ok: true, idDocumento: nuevoId, contenido: bloqueTexto.text, version: version };
}

function generarConvocatoriaYDocumentos(params) {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  const fechaLimite = calcularFechaLimiteAsamblea(ej.fechaCierre);
  const fechaAsamblea = parsearFechaSegura(params.fechaAsamblea);
  const fueraDeTermino = fechaAsamblea > fechaLimite;

  const autRes = listarAutoridades();
  const vencenEsteAnio = autRes.ok ? autRes.vencenEsteAnio : [];

  const ultimoNumero = obtenerUltimoNumeroActa();
  const nuevoNumeroActa = ultimoNumero + 1;

  // Orden del Día -- objeto único, todos los documentos lo citan literal, nunca se retipea
  const puntos = [];
  puntos.push('1º Elección del asociado que presidirá la Asamblea y del Secretario de Asamblea, y de dos socios para firmar el Acta, conforme al artículo 66 del Estatuto Social.');
  puntos.push('2º Lectura y aprobación del Acta de la Asamblea anterior.');
  let n = 3;
  if (fueraDeTermino) {
    puntos.push(n + 'º Consideración de las razones por las cuales la Asamblea General Ordinaria correspondiente al Ejercicio Económico N.° ' + ej.numero + ' se celebra fuera del plazo previsto por el artículo 39 del Estatuto Social.');
    n++;
  }
  puntos.push(n + 'º Lectura y aprobación de la Memoria, Balance General, Inventario, Cuenta de Gastos y Recursos e Informe del Revisor de Cuentas correspondientes al Ejercicio Económico N.° ' + ej.numero + ', finalizado el ' + ej.fechaCierre + '.');
  n++;
  puntos.push(n + 'º Aprobación del monto establecido por la Comisión Directiva para la cuota social, de conformidad con lo establecido en el artículo 16 del Estatuto Social.');
  n++;
  if (vencenEsteAnio.length > 0) {
    puntos.push(n + 'º Renovación parcial de la Comisión Directiva y/o Comisión Revisora de Cuentas por finalización de mandato: ' + vencenEsteAnio.join(', ') + ', conforme a los artículos 22, 35 y 63 del Estatuto Social.');
    n++;
  }
  puntos.push(n + 'º Cualquier otro asunto incluido regularmente en la convocatoria.');

  const ordenDelDiaTexto = puntos.join('\n');

  const nuevoIdAsamblea = 'ASM-' + ej.numero;
  const hojaAsam = getHojaAsambleas();
  // Upsert: si ya hay una fila para este Ejercicio, la actualiza en vez de duplicarla
  const datosAsam = hojaAsam.getDataRange().getValues();
  let filaExistente = -1;
  for (let i = 1; i < datosAsam.length; i++) {
    if (datosAsam[i][1] === ej.idEjercicio) { filaExistente = i + 1; break; }
  }
  if (filaExistente > -1) {
    hojaAsam.getRange(filaExistente, 3).setValue(parsearFechaSegura(params.fechaAsamblea));
    hojaAsam.getRange(filaExistente, 4).setNumberFormat('@').setValue(params.horaAsamblea);
    hojaAsam.getRange(filaExistente, 5).setValue(params.lugarAsamblea);
    hojaAsam.getRange(filaExistente, 6).setValue(JSON.stringify(puntos));
  } else {
    hojaAsam.appendRow([nuevoIdAsamblea, ej.idEjercicio, parsearFechaSegura(params.fechaAsamblea), params.horaAsamblea, params.lugarAsamblea, JSON.stringify(puntos), 'CONVOCADA', '', '', '']);
    hojaAsam.getRange(hojaAsam.getLastRow(), 4).setNumberFormat('@').setValue(params.horaAsamblea);
  }

  const fechaAsambleaFmt = Utilities.formatDate(fechaAsamblea, Session.getScriptTimeZone(), "dd 'de' MMMM 'de' yyyy");

  // ---- 1. Acta de Comisión Directiva ----
  let actaCD = 'ACTA N.º ' + nuevoNumeroActa + '\n\n';
  actaCD += 'En la sede social del Club de Campo La Eugenia, siendo las ' + params.horaReunionCD + ' hs. del día ' + params.fechaReunionCD + ', se reúnen los siguientes miembros de Comisión Directiva: ' + params.presentesCD + '.\n\n';
  actaCD += 'Como primer punto del orden del día se da lectura al acta N.º ' + ultimoNumero + '. Se aprueba por unanimidad.\n\n';
  actaCD += 'Como segundo punto del orden del día, la Comisión Directiva resuelve convocar a Asamblea General Ordinaria para el día ' + fechaAsambleaFmt + ', a las ' + params.horaAsamblea + ' horas, en ' + params.lugarAsamblea + '.\n\n';
  if (fueraDeTermino) {
    actaCD += 'Se deja constancia de que, conforme al artículo 39 del Estatuto Social, el plazo estatutario para la Asamblea Ordinaria vencía el ' + Utilities.formatDate(fechaLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '. En razón de ' + (params.motivoFueraDeTermino || '[COMPLETAR MOTIVO]') + ', se incluye en el Orden del Día el tratamiento de esta circunstancia.\n\n';
  }
  actaCD += 'Como tercer punto del orden del día, se aprueba el siguiente Orden del Día para la Asamblea:\n\n' + ordenDelDiaTexto + '\n\n';
  actaCD += 'Siendo las ' + (params.horaFinReunionCD || '21:15') + ' horas se levanta la sesión.-';

  // ---- 2. Edicto diario local ----
  let edictoDiario = 'CLUB DE CAMPO "LA EUGENIA"\nConvocatoria a Asamblea General Ordinaria\n\n';
  edictoDiario += 'La Comisión Directiva convoca a los señores asociados a la Asamblea General Ordinaria, que se celebrará el día ' + fechaAsambleaFmt + ', a las ' + params.horaAsamblea + ' horas, en ' + params.lugarAsamblea + ', para tratar el siguiente:\n\nORDEN DEL DÍA\n\n' + ordenDelDiaTexto;
  edictoDiario += '\n\nConforme al artículo 44 del Estatuto Social, si a la hora fijada no se reuniera la mayoría absoluta de los asociados con derecho a voto, la Asamblea se celebrará válidamente una hora después.\n\nGarupá, Misiones.';

  // ---- 3. Edicto Boletín Oficial ----
  let edictoBoletin = edictoDiario + '\n\n[Publíquese 2 (dos) días en el Boletín Oficial de la Provincia de Misiones. El plazo de 15 días corridos se cuenta desde la última publicación hasta la Asamblea.]';

  // ---- 4. Circular a socios ----
  let circular = 'Señores Socios:\n\nPor medio de la presente, la Comisión Directiva convoca a la Asamblea General Ordinaria, que se celebrará el día ' + fechaAsambleaFmt + ', a las ' + params.horaAsamblea + ' horas, en ' + params.lugarAsamblea + ', para tratar el siguiente:\n\nORDEN DEL DÍA\n\n' + ordenDelDiaTexto;
  circular += '\n\nRecordamos que, conforme a los artículos 13 y 47 del Estatuto Social, solo podrán participar y votar los socios activos que se encuentren al día con el pago de sus cuotas sociales.\n\nCOMISIÓN DIRECTIVA';

  // Guarda los 4 documentos en la tabla única
  const hoja = getHojaDocumentos();
  const docs = [
    { tipo: 'ACTA_CD_CONVOCATORIA', contenido: actaCD },
    { tipo: 'EDICTO_DIARIO', contenido: edictoDiario },
    { tipo: 'EDICTO_BOLETIN', contenido: edictoBoletin },
    { tipo: 'CIRCULAR', contenido: circular }
  ];
  const resultado = [];
  docs.forEach(doc => {
    const nuevoId = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + doc.tipo;
    hoja.appendRow([nuevoId, ej.idEjercicio, doc.tipo, 1, 'BORRADOR', doc.contenido, 'IA', new Date()]);
    resultado.push({ tipo: doc.tipo, idDocumento: nuevoId, contenido: doc.contenido, estado: 'BORRADOR' });
  });

  return { ok: true, documentos: resultado, fueraDeTermino: fueraDeTermino, fechaLimite: Utilities.formatDate(fechaLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy'), numeroActa: nuevoNumeroActa };
}

function registrarInformeRevisor(params) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'Falta configurar ANTHROPIC_API_KEY en Script Properties' };

  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  const carpeta = DriveApp.getFolderById(CARPETA_INFORMES_REVISOR_ID);
  const archivosIter = carpeta.getFilesByType(MimeType.PDF);
  const archivos = [];
  while (archivosIter.hasNext()) archivos.push(archivosIter.next());
  if (archivos.length === 0) return { ok: false, error: 'No hay ningún PDF en la carpeta de Informes del Revisor. Subí uno primero.' };

  const contentBlocks = archivos.map(f => ({
    type: 'document',
    source: { type: 'base64', media_type: 'application/pdf', data: Utilities.base64Encode(f.getBlob().getBytes()) }
  }));
  contentBlocks.push({ type: 'text', text: 'Extraé el resumen según las instrucciones.' });

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 1000,
    thinking: { type: 'disabled' },
    system: PROMPT_RESUMEN_INFORME_REVISOR,
    messages: [{ role: 'user', content: contentBlocks }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  let resumenInfo = { fechaInforme: '', conclusion: 'INDETERMINADO', resumen: '' };
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    const bloqueTexto = (data.content || []).find(b => b.type === 'text');
    if (bloqueTexto) {
      try {
        resumenInfo = JSON.parse(bloqueTexto.text.replace(/```json|```/g, '').trim());
      } catch (e) { /* deja el resumen vacío si no se pudo parsear, no bloquea el registro */ }
    }
  }

  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  let version = 1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] && datos[i][2] === 'INFORME_REVISOR' && datos[i][1] === ej.idEjercicio) version++;
  }

  const nuevoId = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const urls = archivos.map(f => f.getUrl()).join(' | ');
  const contenidoRegistro = 'ARCHIVO ORIGINAL: ' + urls + '\n\nRESUMEN (generado automáticamente, no reemplaza al documento original):\n' + resumenInfo.resumen + '\n\nConclusión detectada: ' + resumenInfo.conclusion + (resumenInfo.fechaInforme ? ('\nFecha del informe: ' + resumenInfo.fechaInforme) : '');

  hoja.appendRow([nuevoId, ej.idEjercicio, 'INFORME_REVISOR', version, 'RECIBIDO', contenidoRegistro, 'SUBIDO', new Date()]);

  // Archiva los PDF ya registrados, para que la carpeta quede lista para el próximo informe
  const carpetaArchivo = carpeta.createFolder('Registrado_' + nuevoId);
  archivos.forEach(f => { carpeta.removeFile(f); carpetaArchivo.addFile(f); });

  return { ok: true, idDocumento: nuevoId, contenido: contenidoRegistro, version: version, conclusion: resumenInfo.conclusion, archivoUrl: urls };
}


function generarBorradorMemoria(params) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'Falta configurar ANTHROPIC_API_KEY en Script Properties' };

  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  // Guarda los números de socios de esta generación, para sugerirlos la próxima vez
  if (params.sociosActivos || params.sociosAlDia || params.sociosAtraso || params.sociosMorosos) {
    guardarValorPorClave('SOCIOS_' + ej.idEjercicio, JSON.stringify({
      activos: params.sociosActivos || '', alDia: params.sociosAlDia || '',
      atraso: params.sociosAtraso || '', morosos: params.sociosMorosos || ''
    }));
  }

  // Junta las novedades SI/EVALUAR del ejercicio activo
  const novedadesRes = listarNovedades({ idEjercicio: ej.idEjercicio });
  // Solo entran las que ya fueron aprobadas explícitamente (SI) -- EVALUAR significa "todavía sin decidir", no se incluye
  const novedadesRelevantes = (novedadesRes.novedades || []).filter(n => n.considerarMemoria === 'SI');
  const pendientesDeEvaluar = (novedadesRes.novedades || []).filter(n => n.considerarMemoria === 'EVALUAR').length;

  // Trae el último balance del ejercicio activo, si existe
  const balancesRes = listarBalances();
  const balanceDelEjercicio = (balancesRes.balances || []).find(b => b.idEjercicio === ej.idEjercicio);
  let datosBalanceTexto = 'No hay Balance procesado todavía para este Ejercicio.';
  if (balanceDelEjercicio) {
    const detalle = obtenerBalance(balanceDelEjercicio.idBalance);
    if (detalle.ok) {
      const d = detalle.balance.datosExtraidos;
      datosBalanceTexto = 'Activo: $' + d.activoTotal + ', Pasivo: $' + d.pasivoTotal + ', Patrimonio Neto: $' + d.patrimonioNeto + ', Superávit del ejercicio: $' + d.superavitEjercicio;
    }
  }

  const inputTexto = 'EJERCICIO: N.° ' + ej.numero + ', del ' + ej.fechaInicio + ' al ' + ej.fechaCierre + '\n\n' +
    'DATOS DEL BALANCE: ' + datosBalanceTexto + '\n\n' +
    'SOCIOS: Activos ' + (params.sociosActivos || '[no informado]') + ', al día ' + (params.sociosAlDia || '[no informado]') + ', con atraso ' + (params.sociosAtraso || '[no informado]') + ', morosos ' + (params.sociosMorosos || '[no informado]') + '\n\n' +
    'NOVEDADES DEL EJERCICIO:\n' + novedadesRelevantes.map(n => '- [' + n.fecha + '] [' + n.categoria + '] ' + n.titulo + ': ' + n.descripcion + (n.monto ? ' (monto: $' + n.monto + ')' : '')).join('\n');

  const configMemoria = obtenerConfigMemoria();
  const promptCompleto = configMemoria.ok && configMemoria.instrucciones
    ? PROMPT_MEMORIA + '\n\nINSTRUCCIONES ADICIONALES DEL ADMINISTRADOR (seguí estas al pie de la letra, tienen prioridad sobre el resto):\n' + configMemoria.instrucciones
    : PROMPT_MEMORIA;

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 8000,
    thinking: { type: 'disabled' },
    system: promptCompleto,
    messages: [{ role: 'user', content: inputTexto }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return { ok: false, error: 'Error de la API (' + response.getResponseCode() + '): ' + response.getContentText() };
  }

  const data = JSON.parse(response.getContentText());
  const bloqueTexto = (data.content || []).find(b => b.type === 'text');
  if (!bloqueTexto) return { ok: false, error: 'La IA no devolvió texto' };

  // Se fuerza por código, no se confía en que la IA siempre lo incluya en la redacción
  let textoFinal = bloqueTexto.text;
  if (textoFinal.indexOf('Ejercicio Económico N.') === -1 && textoFinal.indexOf('Ejercicio N.') === -1) {
    textoFinal = textoFinal.replace(
      /correspondientes al ejercicio finalizado el/i,
      'correspondientes al Ejercicio Económico N.\u00b0 ' + ej.numero + ', finalizado el'
    );
  }

  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  let version = 1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] && datos[i][2] === 'MEMORIA' && datos[i][1] === ej.idEjercicio) version++;
  }

  const nuevoId = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  hoja.appendRow([nuevoId, ej.idEjercicio, 'MEMORIA', version, 'BORRADOR', textoFinal, 'IA', new Date()]);

  return { ok: true, idDocumento: nuevoId, contenido: textoFinal, novedadesUsadas: novedadesRelevantes.length, version: version, pendientesDeEvaluar: pendientesDeEvaluar };
}

function listarDocumentos(params) {
  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  const tipo = params && params.tipo ? params.tipo : null;
  const idEjercicio = params && params.idEjercicio ? params.idEjercicio : null;

  const documentos = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0]) continue;
    if (tipo && fila[2] !== tipo) continue;
    if (idEjercicio && fila[1] !== idEjercicio) continue;
    documentos.push({ idDocumento: fila[0], idEjercicio: fila[1], tipo: fila[2], version: fila[3], estado: fila[4], contenido: fila[5], generadoPor: fila[6] });
  }
  documentos.sort((a, b) => b.version - a.version);
  return { ok: true, documentos: documentos };
}

function actualizarDocumento(params) {
  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(params.idDocumento)) {
      if (params.contenido) hoja.getRange(i + 1, 6).setValue(params.contenido);
      if (params.estado) hoja.getRange(i + 1, 5).setValue(params.estado);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Documento no encontrado' };
}
