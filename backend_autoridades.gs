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
const ACCIONES_PROTEGIDAS = ['guardarAutoridad', 'eliminarAutoridad', 'guardarNota', 'guardarNotasSeleccionadas', 'eliminarNota', 'generarBorradorActa', 'actualizarActa', 'editarActaFormal', 'asentarActa', 'eliminarBorradorActa', 'anularActa', 'registrarActaManual', 'diagnosticarEliminacionActa', 'eliminarRegistroErroneo', 'migrarActasV2', 'migrarActasV3ModoContenido', 'migrarColumnaPuntosManuales', 'migrarColumnaActaLeida', 'diagnosticarEtiquetadoActas', 'procesarBalance', 'actualizarEstadoBalance', 'cerrarYAbrirNuevoEjercicio', 'actualizarObservacionBalance', 'guardarNovedad', 'actualizarNovedad', 'generarBorradorMemoria', 'registrarInformeRevisor', 'generarConvocatoriaYDocumentos', 'actualizarOrdenDelDiaEnDocumentos', 'guardarPuntosManualesAsamblea', 'actualizarDocumento', 'eliminarDocumento', 'guardarNovedadesSeleccionadas', 'extraerNovedadesDeChat', 'eliminarNovedad', 'guardarConfigMemoria', 'sembrarAsambleaReal', 'guardarAsambleaEjercicio', 'corregirFechaAsambleaExistente', 'limpiarDuplicadosAsambleas', 'backupAsambleas', 'backupDocumentos', 'diagnosticoDuplicadosDocumentos', 'limpiarDocumentosConvocatoriaViejos', 'insertarEncabezadoAsambleas', 'marcarAsambleaCelebrada', 'registrarAutoridadesElectas', 'generarActaAsamblea'];

function doGet(e) {
  const action = e.parameter.action;
  let resultado;

  if (ACCIONES_PROTEGIDAS.indexOf(action) !== -1) {
    if (!validarPinInterno(e.parameter.pin, 'gestion-institucional')) {
      return jsonpResponse({ ok: false, error: 'PIN inválido o sin permiso para este módulo' }, e.parameter.callback);
    }
  }

  try {
    resultado = manejarAccion(action, e.parameter);
  } catch (err) {
    resultado = { ok: false, error: err.message };
  }
  return jsonpResponse(resultado, e.parameter.callback);
}

// CRITERIO TÉCNICO GENERAL DE LA APP: toda acción que envíe contenido textual largo (actas,
// documentos, texto libre, etc.) debe usar POST, nunca GET/JSONP -- una URL armada con
// parámetros largos puede truncarse en el camino (navegador, Apps Script, proxies intermedios),
// y la escritura puede completarse en el servidor mientras la respuesta nunca vuelve al
// cliente (quedando la UI "colgada"), o peor, el contenido puede llegar cortado sin que nada lo
// avise. Ver el módulo de Actas (editarActaFormal, registrarActaManual, generarBorradorActa,
// etc.) para el patrón de referencia.
function doPost(e) {
  let resultado;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (ACCIONES_PROTEGIDAS.indexOf(action) !== -1) {
      if (!validarPinInterno(body.pin, 'gestion-institucional')) {
        return jsonpResponsePost({ ok: false, error: 'PIN inválido o sin permiso para este módulo' });
      }
    }

    resultado = manejarAccion(action, body);
  } catch (err) {
    resultado = { ok: false, error: 'Error interno: ' + err.message };
  }
  return jsonpResponsePost(resultado);
}

function jsonpResponsePost(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Router único, compartido por doGet (lectura y compatibilidad hacia atrás) y doPost (escritura
// de contenido largo). "params" reemplaza a "e.parameter" -- mismo objeto de claves/valores
// venga de donde venga.
function manejarAccion(action, params) {
  let resultado;
    if (action === 'listarAutoridades') {
      resultado = listarAutoridades();
    } else if (action === 'listarAutoridadesTodas') {
      resultado = listarAutoridadesTodas();
    } else if (action === 'eliminarAutoridad') {
      resultado = eliminarAutoridad(params);
    } else if (action === 'guardarAutoridad') {
      resultado = guardarAutoridad(params);
    } else if (action === 'guardarNota') {
      resultado = guardarNota(params);
    } else if (action === 'guardarNotasSeleccionadas') {
      resultado = guardarNotasSeleccionadas(params);
    } else if (action === 'eliminarNota') {
      resultado = eliminarNota(params);
    } else if (action === 'listarNotasPendientes') {
      resultado = listarNotasPendientes();
    } else if (action === 'listarActas') {
      resultado = listarActas();
    } else if (action === 'obtenerActa') {
      resultado = obtenerActa(params.idRegistro);
    } else if (action === 'diagnosticarEliminacionActa') {
      resultado = diagnosticarEliminacionActa(params.idRegistro);
    } else if (action === 'eliminarRegistroErroneo') {
      resultado = eliminarRegistroErroneo(params);
    } else if (action === 'generarBorradorActa') {
      resultado = generarBorradorActa(params);
    } else if (action === 'actualizarActa') {
      resultado = actualizarActa(params);
    } else if (action === 'editarActaFormal') {
      resultado = editarActaFormal(params);
    } else if (action === 'asentarActa') {
      resultado = asentarActa(params);
    } else if (action === 'eliminarBorradorActa') {
      resultado = eliminarBorradorActa(params);
    } else if (action === 'anularActa') {
      resultado = anularActa(params);
    } else if (action === 'registrarActaManual') {
      resultado = registrarActaManual(params);
    } else if (action === 'migrarActasV2') {
      resultado = migrarActasV2();
    } else if (action === 'migrarActasV3ModoContenido') {
      resultado = migrarActasV3ModoContenido();
    } else if (action === 'procesarBalance') {
      resultado = procesarBalance();
    } else if (action === 'listarBalances') {
      resultado = listarBalances();
    } else if (action === 'obtenerBalance') {
      resultado = obtenerBalance(params.idBalance);
    } else if (action === 'actualizarEstadoBalance') {
      resultado = actualizarEstadoBalance(params);
    } else if (action === 'version') {
      resultado = { ok: true, version: 'v59i-revertir-cuota-social-20ago' };
    } else if (action === 'obtenerEjercicioActivo') {
      resultado = obtenerEjercicioActivo();
    } else if (action === 'obtenerResumenDashboard') {
      resultado = obtenerResumenDashboard();
    } else if (action === 'listarEjercicios') {
      resultado = listarEjercicios();
    } else if (action === 'cerrarYAbrirNuevoEjercicio') {
      resultado = cerrarYAbrirNuevoEjercicio(params);
    } else if (action === 'actualizarObservacionBalance') {
      resultado = actualizarObservacionBalance(params);
    } else if (action === 'guardarNovedad') {
      resultado = guardarNovedad(params);
    } else if (action === 'listarNovedades') {
      resultado = listarNovedades(params);
    } else if (action === 'actualizarNovedad') {
      resultado = actualizarNovedad(params);
    } else if (action === 'generarBorradorMemoria') {
      resultado = generarBorradorMemoria(params);
    } else if (action === 'registrarInformeRevisor') {
      resultado = registrarInformeRevisor(params);
    } else if (action === 'generarConvocatoriaYDocumentos') {
      resultado = generarConvocatoriaYDocumentos(params);
    } else if (action === 'actualizarOrdenDelDiaEnDocumentos') {
      resultado = actualizarOrdenDelDiaEnDocumentos(params);
    } else if (action === 'previsualizarOrdenDelDia') {
      resultado = previsualizarOrdenDelDia();
    } else if (action === 'guardarPuntosManualesAsamblea') {
      resultado = guardarPuntosManualesAsamblea(params);
    } else if (action === 'migrarColumnaPuntosManuales') {
      resultado = migrarColumnaPuntosManuales();
    } else if (action === 'diagnosticarEtiquetadoActas') {
      resultado = diagnosticarEtiquetadoActas();
    } else if (action === 'migrarColumnaActaLeida') {
      resultado = migrarColumnaActaLeida();
    } else if (action === 'obtenerAsambleaEjercicio') {
      resultado = obtenerAsambleaEjercicio();
    } else if (action === 'obtenerTextosConvocatoria') {
      resultado = textosConvocatoria();
    } else if (action === 'guardarAsambleaEjercicio') {
      resultado = guardarAsambleaEjercicio(params);
    } else if (action === 'corregirFechaAsambleaExistente') {
      resultado = corregirFechaAsambleaExistente();
    } else if (action === 'limpiarDuplicadosAsambleas') {
      resultado = limpiarDuplicadosAsambleas();
    } else if (action === 'backupAsambleas') {
      resultado = backupAsambleas();
    } else if (action === 'backupDocumentos') {
      resultado = backupDocumentos();
    } else if (action === 'diagnosticoDuplicadosDocumentos') {
      resultado = diagnosticoDuplicadosDocumentos();
    } else if (action === 'limpiarDocumentosConvocatoriaViejos') {
      resultado = limpiarDocumentosConvocatoriaViejos();
    } else if (action === 'insertarEncabezadoAsambleas') {
      resultado = insertarEncabezadoAsambleas();
    } else if (action === 'marcarAsambleaCelebrada') {
      resultado = marcarAsambleaCelebrada(params);
    } else if (action === 'registrarAutoridadesElectas') {
      resultado = registrarAutoridadesElectas(params);
    } else if (action === 'generarActaAsamblea') {
      resultado = generarActaAsamblea(params);
    } else if (action === 'diagnosticoAsambleas') {
      resultado = diagnosticoAsambleas();
    } else if (action === 'sembrarAsambleaReal') {
      resultado = sembrarAsambleaReal();
    } else if (action === 'obtenerConfigMemoria') {
      resultado = obtenerConfigMemoria();
    } else if (action === 'guardarConfigMemoria') {
      resultado = guardarConfigMemoria(params);
    } else if (action === 'obtenerUltimosSocios') {
      resultado = obtenerUltimosSocios(params);
    } else if (action === 'listarDocumentos') {
      resultado = listarDocumentos(params);
    } else if (action === 'eliminarDocumento') {
      resultado = eliminarDocumento(params);
    } else if (action === 'actualizarDocumento') {
      resultado = actualizarDocumento(params);
    } else if (action === 'extraerNovedadesDeChat') {
      resultado = extraerNovedadesDeChat(params);
    } else if (action === 'listarNovedadesDesdeBoletin') {
      resultado = listarNovedadesDesdeBoletin();
    } else if (action === 'guardarNovedadesSeleccionadas') {
      resultado = guardarNovedadesSeleccionadas(params);
    } else if (action === 'eliminarNovedad') {
      resultado = eliminarNovedad(params);
    } else {
      resultado = { ok: false, error: 'Acción no reconocida' };
    }
  return resultado;
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
  const vencenEsteAnioObjs = [];

  filas.forEach(fila => {
    if (!fila[idx.NOMBRE]) return; // fila vacía
    const fechaFin = fila[idx.FECHA_FIN_MANDATO];
    const anioFin = fechaFin ? new Date(fechaFin).getFullYear() : null;
    let vence = 'ok';
    if (anioFin === anioActual) {
      vence = 'este_anio';
      // Solo los órganos con régimen de renovación conocido estatutariamente entran acá --
      // aunque algún día ARQUITECTURA tuviera una FECHA_FIN_MANDATO cargada (a mano, por error),
      // nunca se mete en el texto de "Renovación parcial..." que cita Arts. 22, 35 y 63.
      if (fila[idx.ORGANO] === 'COMISION_DIRECTIVA' || fila[idx.ORGANO] === 'REVISORA_CUENTAS') {
        vencenEsteAnioObjs.push({ cargo: fila[idx.CARGO], nombre: fila[idx.NOMBRE] });
      }
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

  // Orden jerárquico institucional -- no el orden en que se cargaron las filas en la hoja.
  const ORDEN_CARGOS = ['Presidente','Vicepresidente','Secretario','Prosecretario','Tesorero','Protesorero',
    'Vocal Titular 1°','Vocal Titular 2°','Vocal Titular 3°','Vocal Suplente 1°','Vocal Suplente 2°','Vocal Suplente 3°',
    'Revisor Titular','Revisor Suplente'];
  vencenEsteAnioObjs.sort((a, b) => ORDEN_CARGOS.indexOf(a.cargo) - ORDEN_CARGOS.indexOf(b.cargo));
  const vencenEsteAnio = vencenEsteAnioObjs.map(v => v.cargo + ' (' + v.nombre + ')');

  return { ok: true, autoridades: autoridades.filter(a => a.estado === 'VIGENTE'), vencenEsteAnio: vencenEsteAnio };
}

// Igual que listarAutoridades pero sin filtrar por estado -- incluye FINALIZADO.
// Sirve para detectar duplicados o filas cargadas por error y poder eliminarlas.
function listarAutoridadesTodas() {
  const hoja = getHoja();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const autoridades = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.NOMBRE]) continue; // fila vacía
    autoridades.push({
      fila: i + 1, // número de fila real en la hoja, útil para depurar
      idAutoridad: fila[idx.ID_AUTORIDAD],
      organo: fila[idx.ORGANO],
      cargo: fila[idx.CARGO],
      grupoEstatutario: fila[idx.GRUPO_ESTATUTARIO],
      nombre: fila[idx.NOMBRE],
      dni: fila[idx.DNI],
      fechaInicioMandato: fila[idx.FECHA_INICIO_MANDATO] ? Utilities.formatDate(new Date(fila[idx.FECHA_INICIO_MANDATO]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      fechaFinMandato: fila[idx.FECHA_FIN_MANDATO] ? Utilities.formatDate(new Date(fila[idx.FECHA_FIN_MANDATO]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      estado: fila[idx.ESTADO]
    });
  }
  return { ok: true, autoridades: autoridades };
}

// Borra directamente una fila de AUTORIDADES por ID. Pensado para errores de carga
// (fecha mal tipeada, duplicados) -- no para bajas reales de un cargo, eso es guardarAutoridad
// (que cierra el VIGENTE anterior y agrega el reemplazo real).
function eliminarAutoridad(params) {
  if (!params.idAutoridad) return { ok: false, error: 'ID requerido' };
  const hoja = getHoja();
  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_AUTORIDAD]) === String(params.idAutoridad)) {
      hoja.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'No encontrado' };
}

// Grupo estatutario automático por cargo (Art. 21/22 del Estatuto)
const GRUPO_POR_CARGO = {
  'Presidente': 2, 'Secretario': 2, 'Tesorero': 2,
  'Vocal Titular 1°': 2, 'Vocal Titular 2°': 2, 'Vocal Titular 3°': 2,
  'Vicepresidente': 1, 'Prosecretario': 1, 'Protesorero': 1,
  'Vocal Suplente 1°': 1, 'Vocal Suplente 2°': 1, 'Vocal Suplente 3°': 1

};

function calcularFechaFinMandato(organo, cargo, fechaInicio) {
  const inicio = parsearFechaSegura(fechaInicio);
  if (organo === 'REVISORA_CUENTAS') {
    // Renovación anual (Art. 35)
    return new Date(inicio.getFullYear() + 1, inicio.getMonth(), inicio.getDate());
  }
  if (organo === 'COMISION_DIRECTIVA') {
    // Mandato de dos años (Art. 22)
    return new Date(inicio.getFullYear() + 2, inicio.getMonth(), inicio.getDate());
  }
  // Cualquier otro órgano (ej. ARQUITECTURA) -- el Estatuto vigente no le fija un régimen de
  // renovación conocido (el Art. 37 que antes lo hacía fue eliminado: "Artículo 37º.- Eliminado.").
  // No se infiere ninguna duración: queda sin fecha de fin automática hasta que exista una norma
  // vigente que la establezca.
  return null;
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

  hoja.appendRow([nuevoId, organo, cargo, grupo, nombre, dni, parsearFechaSegura(fechaInicio), fechaFin, 'VIGENTE', '']);

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

function guardarNotasSeleccionadas(params) {
  const notas = JSON.parse(params.notas); // [{texto, cargadoPor}]
  const hoja = getHojaBitacora();
  let guardadas = 0;
  notas.forEach(n => {
    const id = 'NOTA-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + guardadas;
    hoja.appendRow([id, new Date(), n.cargadoPor || '', n.texto, false, '']);
    guardadas++;
  });
  return { ok: true, guardadas: guardadas };
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

function getHojaAuditoriaActas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hoja = ss.getSheetByName('AUDITORIA_ACTAS');
  if (!hoja) {
    hoja = ss.insertSheet('AUDITORIA_ACTAS');
    hoja.appendRow(['FECHA', 'ACCION', 'ID_REGISTRO', 'NUMERO_ACTA', 'MOTIVO', 'ORIGEN', 'DATOS_SNAPSHOT']);
    hoja.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#135457').setFontColor('#c4df57');
  }
  return hoja;
}

function logAuditoriaActa_(accion, idRegistro, numeroActa, motivo, origen, snapshot) {
  const hoja = getHojaAuditoriaActas();
  hoja.appendRow([new Date(), accion, idRegistro || '', numeroActa || '', motivo || '', origen || '', snapshot ? JSON.stringify(snapshot) : '']);
}

// Arma y agrega una fila a ACTAS usando nombres de columna, no posiciones fijas -- así el
// orden real de columnas en la hoja (que puede variar tras la migración) nunca desalinea los datos.
function appendFilaActa_(hoja, datosPorNombre) {
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = headers.map(h => (datosPorNombre.hasOwnProperty(h) ? datosPorNombre[h] : ''));
  hoja.appendRow(fila);
  return hoja.getLastRow();
}

function colDeHoja_(hoja, nombre) {
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  return headers.indexOf(nombre) + 1; // 1-indexado; 0 si no existe
}

function listarActas() {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  // Número más alto entre las ASENTADA/ANULADA -- se usa para calcular el "previsto" de cada BORRADOR
  let maxDefinitivo = ULTIMA_ACTA_HISTORICA;
  for (let i = 1; i < datos.length; i++) {
    const n = Number(datos[i][idx.NUMERO_ACTA]);
    const est = datos[i][idx.ESTADO];
    if (!isNaN(n) && (est === 'ASENTADA' || est === 'ANULADA') && n > maxDefinitivo) maxDefinitivo = n;
  }

  const actas = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_REGISTRO]) continue;
    const estado = fila[idx.ESTADO];
    actas.push({
      idRegistro: fila[idx.ID_REGISTRO],
      numeroActa: (estado === 'ASENTADA' || estado === 'ANULADA') ? fila[idx.NUMERO_ACTA] : null,
      numeroPrevisto: estado === 'BORRADOR' ? (maxDefinitivo + 1) : null,
      idEjercicio: fila[idx.ID_EJERCICIO] || '(sin asignar)',
      fechaReunion: fila[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(fila[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      estado: estado,
      origen: fila[idx.ORIGEN] || '',
      motivoAnulacion: fila[idx.MOTIVO_ANULACION] || '',
      modoContenido: fila[idx.MODO_CONTENIDO] || 'ESTRUCTURADO'
    });
  }
  // Ordena: primero por número definitivo descendente, los BORRADOR (sin número) van arriba de todo
  actas.sort((a, b) => {
    if (a.estado === 'BORRADOR' && b.estado !== 'BORRADOR') return -1;
    if (a.estado !== 'BORRADOR' && b.estado === 'BORRADOR') return 1;
    if (a.estado === 'BORRADOR' && b.estado === 'BORRADOR') return 0;
    return Number(b.numeroActa) - Number(a.numeroActa);
  });
  // Dato GLOBAL del módulo, no asociado a ningún borrador individual -- se calcula en vivo,
  // nunca se guarda ni se reserva. Ver regla de UX: un borrador nunca tiene número de Acta.
  return { ok: true, actas: actas, proximoNumeroInstitucional: maxDefinitivo + 1 };
}

// Devuelve el último NÚMERO DE ACTA definitivo (ASENTADA o ANULADA). Los BORRADOR nunca
// participan de este cálculo -- no reservan ni consumen número mientras estén en ese estado.
function obtenerUltimoNumeroActa() {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let maxId = ULTIMA_ACTA_HISTORICA;
  for (let i = 1; i < datos.length; i++) {
    const est = datos[i][idx.ESTADO];
    if (est !== 'ASENTADA' && est !== 'ANULADA') continue;
    const n = Number(datos[i][idx.NUMERO_ACTA]);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return maxId;
}

function obtenerActa(idRegistro) {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_REGISTRO]) === String(idRegistro)) {
      const fila = datos[i];
      return {
        ok: true,
        acta: {
          idRegistro: fila[idx.ID_REGISTRO],
          numeroActa: fila[idx.NUMERO_ACTA] || null,
          numeroActaAnterior: fila[idx.NUMERO_ACTA_ANTERIOR],
          fechaReunion: fila[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(fila[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          horaInicio: formatearHoraSegura(fila[idx.HORA_INICIO]),
          horaFin: formatearHoraSegura(fila[idx.HORA_FIN]),
          presentes: fila[idx.PRESENTES],
          puntos: fila[idx.PUNTOS] ? JSON.parse(fila[idx.PUNTOS]) : [],
          estado: fila[idx.ESTADO],
          origen: fila[idx.ORIGEN] || '',
          motivoAnulacion: fila[idx.MOTIVO_ANULACION] || '',
          modoContenido: fila[idx.MODO_CONTENIDO] || 'ESTRUCTURADO',
          textoLibre: fila[idx.TEXTO_LIBRE] || '',
          actaLeidaNumero: fila[idx.ACTA_LEIDA_NUMERO] || null
        }
      };
    }
  }
  return { ok: false, error: 'Acta no encontrada' };
}

function generarBorradorActa(params) {
  const lock = LockService.getScriptLock();
  let conseguido = false;
  try {
    conseguido = lock.tryLock(10000); // 10 segundos de timeout
    if (!conseguido) {
      return { ok: false, error: 'Otra generación de Acta está en curso ahora mismo. Esperá unos segundos y reintentá.' };
    }

    const hojaActas = getHojaActas();
    const hojaBitacora = getHojaBitacora();

    const idRegistro = 'ACTA-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const ultimoNumero = obtenerUltimoNumeroActa();

    // Punto 1 siempre fijo, encadenado a la última acta ASENTADA real -- tipo propio para poder
    // encontrarlo y actualizarlo después sin depender de su posición ni de parsear texto.
    const puntos = [
      { orden: 1, tipo: 'LECTURA_ACTA_ANTERIOR', texto: 'Se da lectura al acta N.º ' + ultimoNumero + '. Se aprueba por unanimidad.' }
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

    appendFilaActa_(hojaActas, {
      ID_EJERCICIO: (function() { const e = obtenerEjercicioActivo(); return e.ok ? e.ejercicio.idEjercicio : (params.idEjercicio || ''); })(),
      FECHA_REUNION: parsearFechaSegura(params.fechaReunion),
      HORA_INICIO: params.horaInicio || '',
      HORA_FIN: '',
      PRESENTES: params.presentes || '',
      PUNTOS: JSON.stringify(puntos),
      ESTADO: 'BORRADOR',
      NUMERO_ACTA_ANTERIOR: '', // se completa recién al asentar, ahí sabemos cuál es realmente
      NUMERO_ACTA: '', // vacío hasta que se asiente
      ID_REGISTRO: idRegistro,
      ORIGEN: 'APP',
      FECHA_ASENTAMIENTO: '',
      MOTIVO_ANULACION: '',
      MODO_CONTENIDO: 'ESTRUCTURADO', // los borradores generados por el sistema siempre parten de puntos
      TEXTO_LIBRE: '',
      ACTA_LEIDA_NUMERO: ultimoNumero // hecho de la reunión, fijado ahora -- no se recalcula solo después
    });
    const filaNueva = hojaActas.getLastRow();
    hojaActas.getRange(filaNueva, colDeHoja_(hojaActas, 'HORA_INICIO')).setNumberFormat('@').setValue(params.horaInicio || ''); // refuerza texto

    // Marca las notas usadas como procesadas, vinculadas al ID_REGISTRO (no a un número que todavía no existe)
    idsNotasUsadas.forEach(n => {
      hojaBitacora.getRange(n.row, idxBit.PROCESADA + 1).setValue(true);
      hojaBitacora.getRange(n.row, idxBit.ID_REGISTRO_DESTINO + 1).setValue(idRegistro);
    });

    return { ok: true, idRegistro: idRegistro, numeroPrevisto: ultimoNumero + 1, puntos: puntos };
  } finally {
    if (conseguido) lock.releaseLock();
  }
}

// Pasa un BORRADOR a ASENTADA: recién en este momento se calcula y graba el número real,
// nunca antes. Bajo lock para que dos "Asentar" simultáneos no puedan pisarse.
function asentarActa(params) {
  const lock = LockService.getScriptLock();
  let conseguido = false;
  try {
    conseguido = lock.tryLock(10000);
    if (!conseguido) return { ok: false, error: 'Otra operación sobre Actas está en curso. Esperá unos segundos y reintentá.' };

    const hoja = getHojaActas();
    const datos = hoja.getDataRange().getValues();
    const idx = {};
    datos[0].forEach((h, i) => idx[h] = i);

    let fila = -1;
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][idx.ID_REGISTRO]) === String(params.idRegistro)) { fila = i + 1; break; }
    }
    if (fila === -1) return { ok: false, error: 'Acta no encontrada' };
    if (datos[fila - 1][idx.ESTADO] !== 'BORRADOR') return { ok: false, error: 'Esta acta ya no está en Borrador -- no se puede volver a asentar.' };

    const numero = obtenerUltimoNumeroActa() + 1;
    const numeroAnteriorCalculado = numero - 1;
    const actaLeidaRaw = datos[fila - 1][idx.ACTA_LEIDA_NUMERO];
    const actaLeida = (actaLeidaRaw !== '' && actaLeidaRaw !== undefined && actaLeidaRaw !== null) ? Number(actaLeidaRaw) : null;

    // ¿Viene una confirmación de una discrepancia ya mostrada al usuario? Se exige que los tres
    // valores confirmados coincidan EXACTO con lo recién recalculado bajo lock -- si algo cambió
    // en el medio (otra acta se asentó/registró), la confirmación vieja queda invalidada y se
    // vuelve a pedir, nunca se asienta con datos desactualizados.
    const esConfirmacion = params.confirmadoDiscrepancia === 'true' || params.confirmadoDiscrepancia === true;
    const situacionCoincide = esConfirmacion &&
      Number(params.numeroAConfirmar) === numero &&
      Number(params.numeroAnteriorAConfirmar) === numeroAnteriorCalculado &&
      (params.actaLeidaAConfirmar === '' || params.actaLeidaAConfirmar === undefined
        ? actaLeida === null
        : Number(params.actaLeidaAConfirmar) === actaLeida);

    const hayDiscrepancia = actaLeida === null || actaLeida !== numeroAnteriorCalculado;

    if (hayDiscrepancia && !situacionCoincide) {
      // Arma el detalle de actas intermedias reales entre lo leído y lo que va a quedar como
      // "acta anterior" -- solo si actaLeida tiene valor (si es null, es el caso "sin dato", no hay
      // rango que armar).
      let actasIntermedias = [];
      if (actaLeida !== null) {
        for (let i = 1; i < datos.length; i++) {
          const n = Number(datos[i][idx.NUMERO_ACTA]);
          const est = datos[i][idx.ESTADO];
          if (!isNaN(n) && (est === 'ASENTADA' || est === 'ANULADA') && n > actaLeida && n <= numeroAnteriorCalculado) {
            actasIntermedias.push({
              numeroActa: n,
              fechaReunion: datos[i][idx.FECHA_REUNION] ? Utilities.formatDate(new Date(datos[i][idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
              origen: datos[i][idx.ORIGEN] || '',
              estado: est
            });
          }
        }
        actasIntermedias.sort((a, b) => a.numeroActa - b.numeroActa);
      }

      return {
        ok: false,
        requiereConfirmacion: true,
        numeroPrevisto: numero,
        numeroAnteriorCalculado: numeroAnteriorCalculado,
        actaLeidaRegistrada: actaLeida,
        actasIntermedias: actasIntermedias
      };
    }

    // Defensivo: revalida que ese número no exista ya en ninguna fila (no debería pasar con el lock, pero por las dudas)
    for (let i = 1; i < datos.length; i++) {
      if (Number(datos[i][idx.NUMERO_ACTA]) === numero) {
        return { ok: false, error: 'Colisión de numeración detectada (N.º ' + numero + ' ya existe). Avisale a Augusto antes de reintentar.' };
      }
    }

    // Ahora que hay número real, se puede completar la referencia "acta anterior" correctamente
    hoja.getRange(fila, idx.NUMERO_ACTA_ANTERIOR + 1).setValue(String(numeroAnteriorCalculado));
    hoja.getRange(fila, idx.NUMERO_ACTA + 1).setValue(numero);
    hoja.getRange(fila, idx.ESTADO + 1).setValue('ASENTADA');
    hoja.getRange(fila, idx.FECHA_ASENTAMIENTO + 1).setValue(new Date());

    logAuditoriaActa_('ASENTAR', params.idRegistro, numero, '', 'APP', { actaLeidaRegistrada: actaLeida, numeroAnteriorCalculado: numeroAnteriorCalculado, confirmoDiscrepancia: esConfirmacion });

    return { ok: true, numeroActa: numero };
  } finally {
    if (conseguido) lock.releaseLock();
  }
}

// Elimina físicamente un BORRADOR (nunca una acta ASENTADA o ANULADA). Como un borrador nunca
// tuvo número real, no deja hueco ni afecta la numeración. Libera las notas de Bitácora que
// había tomado, para que vuelvan a estar disponibles para la próxima acta.
// SOLO LECTURA -- no borra, no edita, no anula. Reúne todas las dependencias reales de un
// registro de ACTAS antes de decidir si (y cómo) se puede eliminar como error de la app.
// No asume nada: si algo no se puede determinar con certeza, se informa como tal, nunca se omite.
function diagnosticarEliminacionActa(idRegistro) {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  let filaData = null;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_REGISTRO]) === String(idRegistro)) { filaData = datos[i]; break; }
  }
  if (!filaData) return { ok: false, error: 'Acta no encontrada' };

  const numeroActa = filaData[idx.NUMERO_ACTA];
  const estado = filaData[idx.ESTADO];

  const snapshot = {
    idRegistro: idRegistro,
    numeroActa: numeroActa || null,
    numeroActaAnterior: filaData[idx.NUMERO_ACTA_ANTERIOR] || null,
    estado: estado,
    origen: filaData[idx.ORIGEN] || '',
    modoContenido: filaData[idx.MODO_CONTENIDO] || 'ESTRUCTURADO',
    fechaReunion: filaData[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(filaData[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
    horaInicio: filaData[idx.HORA_INICIO] || '',
    horaFin: filaData[idx.HORA_FIN] || '',
    presentes: filaData[idx.PRESENTES] || '',
    puntos: filaData[idx.PUNTOS] || '[]',
    textoLibre: filaData[idx.TEXTO_LIBRE] || ''
  };

  // 1. Notas de Bitácora vinculadas
  const hojaBitacora = getHojaBitacora();
  const datosBit = hojaBitacora.getDataRange().getValues();
  const idxBit = {};
  datosBit[0].forEach((h, i) => idxBit[h] = i);
  const notasBitacoraVinculadas = [];
  for (let i = 1; i < datosBit.length; i++) {
    if (String(datosBit[i][idxBit.ID_REGISTRO_DESTINO]) === String(idRegistro)) {
      notasBitacoraVinculadas.push({ idNota: datosBit[i][idxBit.ID_NOTA], texto: datosBit[i][idxBit.TEXTO] });
    }
  }

  // 2. Otras actas que la citan como "acta anterior" (referencia narrativa, no estructural)
  const actasQueLaReferencian = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][idx.ID_REGISTRO] === idRegistro) continue;
    if (numeroActa && String(datos[i][idx.NUMERO_ACTA_ANTERIOR]) === String(numeroActa)) {
      actasQueLaReferencian.push({ idRegistro: datos[i][idx.ID_REGISTRO], numeroActa: datos[i][idx.NUMERO_ACTA], estado: datos[i][idx.ESTADO] });
    }
  }

  // 3. Vínculo con Asamblea (Acta de Comisión Directiva de Convocatoria)
  let vinculadaAAsamblea = null;
  const hojaAsam = getHojaAsambleas();
  const headersAsam = hojaAsam.getRange(1, 1, 1, hojaAsam.getLastColumn()).getValues()[0];
  const colIdActaCD = headersAsam.indexOf('ID_REGISTRO_ACTA_CD');
  if (colIdActaCD !== -1) {
    const datosAsam = hojaAsam.getDataRange().getValues();
    for (let i = 1; i < datosAsam.length; i++) {
      if (String(datosAsam[i][colIdActaCD]) === String(idRegistro)) {
        vinculadaAAsamblea = { idEjercicio: datosAsam[i][1], idAsamblea: datosAsam[i][0] };
        break;
      }
    }
  }

  // 4. Documentos (Convocatoria/Edictos/Circular/etc.) que mencionan literalmente este número
  const documentosConMencion = [];
  if (numeroActa) {
    const hojaDocs = getHojaDocumentos();
    const datosDocs = hojaDocs.getDataRange().getValues();
    const idxDocs = {};
    datosDocs[0].forEach((h, i) => idxDocs[h] = i);
    const patronFuerte = new RegExp('acta\\s*n\\.?°?º?\\s*' + numeroActa + '\\b', 'i');
    const patronDebil = new RegExp('\\b' + numeroActa + '\\b');
    for (let i = 1; i < datosDocs.length; i++) {
      const contenido = String(datosDocs[i][idxDocs.CONTENIDO] || '');
      if (!contenido) continue;
      const matchFuerte = patronFuerte.test(contenido);
      const matchDebil = !matchFuerte && patronDebil.test(contenido);
      if (matchFuerte || matchDebil) {
        documentosConMencion.push({
          idDocumento: datosDocs[i][idxDocs.ID_DOCUMENTO],
          tipo: datosDocs[i][idxDocs.TIPO],
          version: datosDocs[i][idxDocs.VERSION],
          estado: datosDocs[i][idxDocs.ESTADO],
          tipoDeMencion: matchFuerte ? 'EXPLICITA (dice "acta N.º ' + numeroActa + '")' : 'POSIBLE (aparece el número, sin contexto de "acta")'
        });
      }
    }
  }

  return {
    ok: true,
    acta: snapshot,
    dependencias: {
      notasBitacoraVinculadas: notasBitacoraVinculadas,
      actasQueLaReferencian: actasQueLaReferencian,
      vinculadaAAsamblea: vinculadaAAsamblea,
      documentosConMencion: documentosConMencion
    },
    resumen: {
      totalNotasBitacora: notasBitacoraVinculadas.length,
      totalActasQueReferencian: actasQueLaReferencian.length,
      tieneVinculoAsamblea: vinculadaAAsamblea !== null,
      totalDocumentosConMencion: documentosConMencion.length
    }
  };
}

// Elimina físicamente una acta ASENTADA/ANULADA que NUNCA existió en el Libro físico -- es
// distinta de "Anular" (que es para actas reales que hay que dejar sin efecto, conservando su
// número para siempre). Acá el registro es un error puro de la app, así que su número queda
// libre de nuevo para la numeración real. Requiere PIN re-ingresado y motivo obligatorio,
// vuelve a correr el diagnóstico completo bajo lock (nunca confía en un diagnóstico previo),
// y aborta sin tocar nada si aparece alguna otra acta que la referencia como anterior.
function eliminarRegistroErroneo(params) {
  if (!params.motivo || !params.motivo.trim()) {
    return { ok: false, error: 'El motivo es obligatorio para eliminar un registro erróneo.' };
  }
  if (!validarPinInterno(params.pin, 'gestion-institucional')) {
    return { ok: false, error: 'PIN inválido. Volvé a ingresarlo para confirmar esta operación.' };
  }

  const lock = LockService.getScriptLock();
  let conseguido = false;
  try {
    conseguido = lock.tryLock(10000);
    if (!conseguido) return { ok: false, error: 'Otra operación sobre Actas está en curso. Esperá unos segundos y reintentá.' };

    // Re-diagnostica desde cero bajo el lock -- nunca confía en un diagnóstico corrido antes
    // de tomar el lock, por si algo cambió en el medio.
    const diag = diagnosticarEliminacionActa(params.idRegistro);
    if (!diag.ok) return diag;

    if (diag.acta.estado === 'BORRADOR') {
      return { ok: false, error: 'Esta acta está en Borrador -- usá "Eliminar borrador", no esta acción (que es para registros ya Asentados/Anulados).' };
    }

    if (diag.dependencias.actasQueLaReferencian.length > 0) {
      return {
        ok: false,
        error: 'No se eliminó nada: hay ' + diag.dependencias.actasQueLaReferencian.length + ' acta(s) que referencian a esta como "acta anterior". Resolvé esa referencia primero (editando esa acta desde "Editar acta asentada/anulada") y volvé a intentar.',
        actasQueLaReferencian: diag.dependencias.actasQueLaReferencian
      };
    }

    const hoja = getHojaActas();
    const datos = hoja.getDataRange().getValues();
    const idx = {};
    datos[0].forEach((h, i) => idx[h] = i);

    let fila = -1;
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][idx.ID_REGISTRO]) === String(params.idRegistro)) { fila = i + 1; break; }
    }
    if (fila === -1) return { ok: false, error: 'Acta no encontrada (puede que ya se haya eliminado).' };

    // Libera las notas de Bitácora vinculadas -- vuelven a estar disponibles para otra acta.
    const hojaBitacora = getHojaBitacora();
    const datosBit = hojaBitacora.getDataRange().getValues();
    const idxBit = {};
    datosBit[0].forEach((h, i) => idxBit[h] = i);
    let notasLiberadas = 0;
    for (let i = 1; i < datosBit.length; i++) {
      if (String(datosBit[i][idxBit.ID_REGISTRO_DESTINO]) === String(params.idRegistro)) {
        hojaBitacora.getRange(i + 1, idxBit.PROCESADA + 1).setValue(false);
        hojaBitacora.getRange(i + 1, idxBit.ID_REGISTRO_DESTINO + 1).setValue('');
        notasLiberadas++;
      }
    }

    // Snapshot completo ANTES de borrar la fila -- incluye el diagnóstico completo, no solo el acta.
    logAuditoriaActa_('ELIMINAR_REGISTRO_ERRONEO', params.idRegistro, diag.acta.numeroActa, params.motivo, diag.acta.estado, {
      acta: diag.acta,
      dependencias: diag.dependencias
    });

    hoja.deleteRow(fila);

    const proximoNumeroDisponible = obtenerUltimoNumeroActa() + 1;

    return {
      ok: true,
      mensaje: 'Registro erróneo eliminado.',
      notasLiberadas: notasLiberadas,
      proximoNumeroDisponible: proximoNumeroDisponible,
      // Estos documentos NO se tocaron -- quedan como advertencia para corregir a mano
      // (ver corrección pendiente del flujo de Convocatoria).
      documentosConMencionSinCorregir: diag.dependencias.documentosConMencion
    };
  } finally {
    if (conseguido) lock.releaseLock();
  }
}

function eliminarBorradorActa(params) {

  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  let fila = -1, filaData = null;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_REGISTRO]) === String(params.idRegistro)) { fila = i + 1; filaData = datos[i]; break; }
  }
  if (fila === -1) return { ok: false, error: 'Acta no encontrada' };
  if (filaData[idx.ESTADO] !== 'BORRADOR') {
    return { ok: false, error: 'Solo se pueden eliminar actas en Borrador. Esta ya está ' + filaData[idx.ESTADO].toLowerCase() + ' -- si hay un error, hay que anularla, no se puede borrar.' };
  }
  if (!params.motivo || !params.motivo.trim()) {
    return { ok: false, error: 'El motivo es obligatorio para eliminar un borrador.' };
  }

  // Libera las notas de Bitácora que este borrador había tomado
  const hojaBitacora = getHojaBitacora();
  const datosBit = hojaBitacora.getDataRange().getValues();
  const idxBit = {};
  datosBit[0].forEach((h, i) => idxBit[h] = i);
  let notasLiberadas = 0;
  for (let i = 1; i < datosBit.length; i++) {
    if (String(datosBit[i][idxBit.ID_REGISTRO_DESTINO]) === String(params.idRegistro)) {
      hojaBitacora.getRange(i + 1, idxBit.PROCESADA + 1).setValue(false);
      hojaBitacora.getRange(i + 1, idxBit.ID_REGISTRO_DESTINO + 1).setValue('');
      notasLiberadas++;
    }
  }

  const snapshot = {
    idRegistro: filaData[idx.ID_REGISTRO],
    fechaReunion: filaData[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(filaData[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
    presentes: filaData[idx.PRESENTES],
    puntos: filaData[idx.PUNTOS]
  };
  logAuditoriaActa_('ELIMINAR_BORRADOR', params.idRegistro, '', params.motivo, 'APP', snapshot);

  hoja.deleteRow(fila);
  return { ok: true, notasLiberadas: notasLiberadas };
}

// Anula una acta ya ASENTADA. Estado terminal: nunca vuelve a Borrador ni a Asentada, nunca se
// borra, y su número NUNCA se reutiliza -- queda "quemado" para siempre, como en un libro
// rubricado físico. Las notas de Bitácora que había tomado NO se liberan (la reunión existió).
function anularActa(params) {
  if (!params.motivo || !params.motivo.trim()) {
    return { ok: false, error: 'El motivo es obligatorio para anular un acta.' };
  }
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_REGISTRO]) === String(params.idRegistro)) {
      const estado = datos[i][idx.ESTADO];
      if (estado === 'BORRADOR') return { ok: false, error: 'No se anula un borrador -- se elimina.' };
      if (estado === 'ANULADA') return { ok: false, error: 'Esta acta ya está anulada.' };

      const numero = datos[i][idx.NUMERO_ACTA];
      hoja.getRange(i + 1, idx.ESTADO + 1).setValue('ANULADA');
      hoja.getRange(i + 1, idx.MOTIVO_ANULACION + 1).setValue(params.motivo);
      logAuditoriaActa_('ANULAR', params.idRegistro, numero, params.motivo, 'APP', null);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Acta no encontrada' };
}

// Registra un acta confeccionada fuera de la app (a mano, en el Libro), con un número indicado
// explícitamente -- puede ser para completar un hueco, no necesariamente el siguiente en la
// secuencia. Rechaza si el número ya existe (en cualquier estado).
function registrarActaManual(params) {
  const numero = Number(params.numeroActa);
  if (!numero || isNaN(numero) || numero <= ULTIMA_ACTA_HISTORICA) {
    return { ok: false, error: 'Número de Acta inválido (tiene que ser mayor a ' + ULTIMA_ACTA_HISTORICA + ').' };
  }

  const modoContenido = params.modoContenido === 'TEXTO_LIBRE' ? 'TEXTO_LIBRE' : 'ESTRUCTURADO';
  if (modoContenido === 'TEXTO_LIBRE' && (!params.textoLibre || !params.textoLibre.trim())) {
    return { ok: false, error: 'Falta el texto completo del acta.' };
  }
  if (modoContenido === 'ESTRUCTURADO' && (!params.puntos || params.puntos === '[]')) {
    return { ok: false, error: 'Cargá al menos un punto tratado.' };
  }

  // Lock compartido con asentarActa/generarBorradorActa: esta función también puede escribir
  // un NUMERO_ACTA definitivo, y el chequeo de duplicados de abajo tiene que ser atómico con
  // esa escritura para que dos registros simultáneos no puedan crear el mismo número.
  const lock = LockService.getScriptLock();
  let conseguido = false;
  try {
    conseguido = lock.tryLock(10000);
    if (!conseguido) return { ok: false, error: 'Otra operación sobre Actas está en curso. Esperá unos segundos y reintentá.' };

    const hoja = getHojaActas();
    const datos = hoja.getDataRange().getValues();
    const idx = {};
    datos[0].forEach((h, i) => idx[h] = i);

    for (let i = 1; i < datos.length; i++) {
      if (Number(datos[i][idx.NUMERO_ACTA]) === numero) {
        return { ok: false, error: 'Ya existe un Acta N.º ' + numero + ' en el sistema (estado: ' + datos[i][idx.ESTADO] + ').' };
      }
    }

    const idRegistro = 'ACTA-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    appendFilaActa_(hoja, {
      ID_EJERCICIO: params.idEjercicio || (function() { const e = obtenerEjercicioActivo(); return e.ok ? e.ejercicio.idEjercicio : ''; })(),
      FECHA_REUNION: parsearFechaSegura(params.fechaReunion),
      HORA_INICIO: params.horaInicio || '',
      HORA_FIN: params.horaFin || '',
      PRESENTES: params.presentes || '',
      PUNTOS: modoContenido === 'ESTRUCTURADO' ? (params.puntos || '[]') : '[]',
      ESTADO: 'ASENTADA',
      NUMERO_ACTA_ANTERIOR: params.numeroActaAnterior || String(numero - 1),
      NUMERO_ACTA: numero,
      ID_REGISTRO: idRegistro,
      ORIGEN: 'MANUAL',
      FECHA_ASENTAMIENTO: params.fechaAsentamiento ? parsearFechaSegura(params.fechaAsentamiento) : new Date(),
      MOTIVO_ANULACION: '',
      MODO_CONTENIDO: modoContenido,
      TEXTO_LIBRE: modoContenido === 'TEXTO_LIBRE' ? params.textoLibre : '',
      // Si no se indica qué acta se leyó realmente en la reunión, se asume igual a la
      // correlatividad calculada -- comportamiento de siempre, sin cambios para quien no lo necesite.
      ACTA_LEIDA_NUMERO: params.actaLeidaNumero || params.numeroActaAnterior || String(numero - 1)
    });
    const fila = hoja.getLastRow();
    hoja.getRange(fila, colDeHoja_(hoja, 'HORA_INICIO')).setNumberFormat('@').setValue(params.horaInicio || '');
    hoja.getRange(fila, colDeHoja_(hoja, 'HORA_FIN')).setNumberFormat('@').setValue(params.horaFin || '');

    logAuditoriaActa_('REGISTRAR_MANUAL', idRegistro, numero, 'Carga de acta manual/externa (' + modoContenido + ')', 'MANUAL', null);

    return { ok: true, idRegistro: idRegistro };
  } finally {
    if (conseguido) lock.releaseLock();
  }
}

// Solo edita el CONTENIDO de un acta todavía en Borrador. Una vez Asentada o Anulada, el
// contenido no se toca más -- si hay un error hay que anular y generar una nueva.
function actualizarActa(params) {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idx.ID_REGISTRO]) === String(params.idRegistro)) {
      if (datos[i][idx.ESTADO] !== 'BORRADOR') {
        return { ok: false, error: 'Solo se puede editar el contenido mientras el acta está en Borrador.' };
      }
      const row = i + 1;
      const modo = datos[i][idx.MODO_CONTENIDO] || 'ESTRUCTURADO';
      let advertenciaTextoInconsistente = null;

      // ACTA_LEIDA_NUMERO es el dato fuente -- si cambia en una acta ESTRUCTURADO, el Punto 1
      // (identificado por tipo, nunca por posición ni por texto) se reescribe a partir de él.
      // Nunca al revés: el texto del punto nunca decide el valor del campo.
      if (params.actaLeidaNumero !== undefined && params.actaLeidaNumero !== '') {
        const nuevoActaLeida = Number(params.actaLeidaNumero);
        if (!nuevoActaLeida || isNaN(nuevoActaLeida)) {
          return { ok: false, error: 'Número de Acta leída inválido.' };
        }

        if (modo === 'ESTRUCTURADO') {
          let puntosArray;
          try {
            puntosArray = params.puntos ? JSON.parse(params.puntos) : JSON.parse(datos[i][idx.PUNTOS] || '[]');
          } catch (e) {
            return { ok: false, error: 'Los puntos actuales tienen un formato inválido -- no se hizo ningún cambio.' };
          }
          const puntosLectura = puntosArray.filter(p => p && p.tipo === 'LECTURA_ACTA_ANTERIOR');
          if (puntosLectura.length !== 1) {
            return { ok: false, error: 'No se pudo actualizar el Punto 1 automáticamente: se esperaba exactamente un punto con tipo LECTURA_ACTA_ANTERIOR y se encontraron ' + puntosLectura.length + '. Revisá la estructura de puntos antes de cambiar el Acta leída. No se guardó ningún cambio.' };
          }
          puntosLectura[0].texto = 'Se da lectura al acta N.º ' + nuevoActaLeida + '. Se aprueba por unanimidad.';
          hoja.getRange(row, idx.PUNTOS + 1).setValue(JSON.stringify(puntosArray));
        } else {
          // TEXTO_LIBRE: nunca se reescribe la prosa -- solo se guarda el dato estructurado, y se
          // hace un chequeo suave (informativo, no bloqueante) contra el texto vigente o el nuevo.
          if (params.puntos) hoja.getRange(row, idx.PUNTOS + 1).setValue(params.puntos); // no debería llegar en este modo, pero por consistencia
          const textoAComparar = params.textoLibre !== undefined ? params.textoLibre : String(datos[i][idx.TEXTO_LIBRE] || '');
          const match = textoAComparar.match(/acta\s*n\.?°?º?\s*(\d+)/i);
          if (match && Number(match[1]) !== nuevoActaLeida) {
            advertenciaTextoInconsistente = 'El texto menciona el Acta N.º ' + match[1] + ' pero el campo estructurado que acabás de guardar dice ' + nuevoActaLeida + '. Revisá cuál es el correcto -- esto no bloquea el guardado.';
          }
        }
        hoja.getRange(row, idx.ACTA_LEIDA_NUMERO + 1).setValue(nuevoActaLeida);
      } else if (params.puntos) {
        // Sin cambio de Acta leída -- comportamiento de siempre, guarda los puntos tal cual vienen.
        hoja.getRange(row, idx.PUNTOS + 1).setValue(params.puntos);
      }

      if (params.horaFin) hoja.getRange(row, idx.HORA_FIN + 1).setNumberFormat('@').setValue(params.horaFin);
      if (params.presentes) hoja.getRange(row, idx.PRESENTES + 1).setValue(params.presentes);

      return { ok: true, advertenciaTextoInconsistente: advertenciaTextoInconsistente };
    }
  }
  return { ok: false, error: 'Acta no encontrada' };
}

// Edita una acta YA ASENTADA o ANULADA (nunca un Borrador -- para eso está actualizarActa).
// El Libro físico es la fuente formal definitiva; esto existe para reconciliar la app con lo
// que finalmente quedó asentado ahí (reordenamientos, correcciones de redacción, fecha, etc.).
// Requiere PIN re-ingresado (no alcanza con la sesión ya abierta) y motivo obligatorio.
// Nunca cambia el ESTADO -- eso sigue siendo exclusivo de asentarActa/anularActa.
function editarActaFormal(params) {
  if (!params.motivo || !params.motivo.trim()) {
    return { ok: false, error: 'El motivo es obligatorio para editar un acta ya asentada o anulada.' };
  }
  // Re-valida el PIN explícitamente en esta acción puntual, además de la protección general
  // de ACCIONES_PROTEGIDAS -- es la "segunda confirmación consciente" pedida para esta acción sensible.
  if (!validarPinInterno(params.pin, 'gestion-institucional')) {
    return { ok: false, error: 'PIN inválido. Volvé a ingresarlo para confirmar esta edición.' };
  }

  // Lock compartido con asentarActa/registrarActaManual/generarBorradorActa: esta función
  // también puede escribir un NUMERO_ACTA nuevo (renumeración), y el chequeo de duplicados
  // tiene que ser atómico con esa escritura.
  const lock = LockService.getScriptLock();
  let conseguido = false;
  try {
    conseguido = lock.tryLock(10000);
    if (!conseguido) return { ok: false, error: 'Otra operación sobre Actas está en curso. Esperá unos segundos y reintentá.' };

    const hoja = getHojaActas();
    const datos = hoja.getDataRange().getValues();
    const idx = {};
    datos[0].forEach((h, i) => idx[h] = i);

    let fila = -1;
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][idx.ID_REGISTRO]) === String(params.idRegistro)) { fila = i + 1; break; }
    }
    if (fila === -1) return { ok: false, error: 'Acta no encontrada' };

    const filaData = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
    const estadoActual = filaData[idx.ESTADO];
    if (estadoActual === 'BORRADOR') {
      return { ok: false, error: 'Esta acta está en Borrador -- usá la edición normal, no esta.' };
    }

    const numeroActual = filaData[idx.NUMERO_ACTA];
    let numeroNuevo = null;

    // Validación absoluta de duplicados: un número no puede repetirse nunca, sin importar el
    // estado de quien ya lo tenga (tampoco se libera automáticamente el de una ANULADA).
    if (params.numeroActa !== undefined && params.numeroActa !== '' && String(params.numeroActa) !== String(numeroActual)) {
      numeroNuevo = Number(params.numeroActa);
      if (!numeroNuevo || isNaN(numeroNuevo) || numeroNuevo <= ULTIMA_ACTA_HISTORICA) {
        return { ok: false, error: 'Número de Acta inválido (tiene que ser mayor a ' + ULTIMA_ACTA_HISTORICA + ').' };
      }
      for (let i = 1; i < datos.length; i++) {
        if (i === fila - 1) continue;
        if (Number(datos[i][idx.NUMERO_ACTA]) === numeroNuevo) {
          return { ok: false, error: 'Ya existe un Acta N.º ' + numeroNuevo + ' en el sistema (estado: ' + datos[i][idx.ESTADO] + '). No se hizo ningún cambio.' };
        }
      }
    }

    // MODO_CONTENIDO: si no viene en params, se conserva el modo actual (una edición de metadata
    // -- fecha, horarios, número -- no tiene por qué forzar ni tocar el modo de contenido).
    const modoActual = filaData[idx.MODO_CONTENIDO] || 'ESTRUCTURADO';
    const modoNuevo = params.modoContenido === 'TEXTO_LIBRE' || params.modoContenido === 'ESTRUCTURADO' ? params.modoContenido : modoActual;
    const huboConversionModo = modoNuevo !== modoActual;

    // Validación obligatoria de contenido al convertir de modo -- si vas a convertir, tiene que
    // venir el contenido correspondiente al modo nuevo. Se rechaza TODO el cambio (nada se
    // escribe) antes que dejar el acta con un modo declarado y sin contenido real adentro.
    if (huboConversionModo) {
      if (modoNuevo === 'TEXTO_LIBRE' && (!params.textoLibre || !params.textoLibre.trim())) {
        return { ok: false, error: 'Para convertir a Texto libre, tenés que pegar el texto completo del acta. No se hizo ningún cambio.' };
      }
      if (modoNuevo === 'ESTRUCTURADO') {
        let puntosValidos = [];
        try { puntosValidos = params.puntos ? JSON.parse(params.puntos) : []; } catch (e) { puntosValidos = []; }
        if (!Array.isArray(puntosValidos) || puntosValidos.filter(p => p && p.texto && String(p.texto).trim()).length === 0) {
          return { ok: false, error: 'Para convertir a Estructurado, tenés que cargar al menos un punto. No se hizo ningún cambio.' };
        }
      }
    }
    // Si NO hay conversión y el modo actual es TEXTO_LIBRE, tampoco se puede guardar vacío
    // por accidente (por ejemplo, si alguien borra todo el texto sin querer).
    if (!huboConversionModo && modoNuevo === 'TEXTO_LIBRE' && params.textoLibre !== undefined && !params.textoLibre.trim()) {
      return { ok: false, error: 'El texto del acta no puede quedar vacío. No se hizo ningún cambio.' };
    }

    // Snapshot ANTES de tocar nada -- incluye modo y contenido de los dos tipos, sin importar cuál esté activo
    const snapshotAntes = {
      numeroActa: filaData[idx.NUMERO_ACTA],
      numeroActaAnterior: filaData[idx.NUMERO_ACTA_ANTERIOR],
      fechaReunion: filaData[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(filaData[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      horaInicio: filaData[idx.HORA_INICIO],
      horaFin: filaData[idx.HORA_FIN],
      presentes: filaData[idx.PRESENTES],
      modoContenido: modoActual,
      puntos: filaData[idx.PUNTOS],
      textoLibre: filaData[idx.TEXTO_LIBRE] || ''
    };

    // Campos de metadata: se distingue "no vino en el POST" (undefined -> no tocar) de
    // "vino explícitamente vacío" (string vacío -> sí escribir el vaciado). El frontend manda
    // SIEMPRE todas las claves en el body POST, así que un '' real significa "vaciar a propósito".
    if (numeroNuevo !== null) hoja.getRange(fila, idx.NUMERO_ACTA + 1).setValue(numeroNuevo);
    if (params.numeroActaAnterior !== undefined) hoja.getRange(fila, idx.NUMERO_ACTA_ANTERIOR + 1).setValue(params.numeroActaAnterior);
    if (params.fechaReunion !== undefined && params.fechaReunion !== '') hoja.getRange(fila, idx.FECHA_REUNION + 1).setValue(parsearFechaSegura(params.fechaReunion));
    if (params.horaInicio !== undefined) hoja.getRange(fila, idx.HORA_INICIO + 1).setNumberFormat('@').setValue(params.horaInicio);
    if (params.horaFin !== undefined) hoja.getRange(fila, idx.HORA_FIN + 1).setNumberFormat('@').setValue(params.horaFin);
    if (params.presentes !== undefined) hoja.getRange(fila, idx.PRESENTES + 1).setValue(params.presentes);

    // Contenido: se escribe SOLO lo correspondiente al modo final, nunca los dos a la vez, y NUNCA
    // se reconstruye ni se traduce entre modos -- lo que venga en params es lo que se guarda, tal cual.
    hoja.getRange(fila, idx.MODO_CONTENIDO + 1).setValue(modoNuevo);
    if (modoNuevo === 'ESTRUCTURADO') {
      if (params.puntos !== undefined) hoja.getRange(fila, idx.PUNTOS + 1).setValue(params.puntos);
      if (huboConversionModo) hoja.getRange(fila, idx.TEXTO_LIBRE + 1).setValue(''); // se limpia el otro modo al convertir
    } else { // TEXTO_LIBRE
      if (params.textoLibre !== undefined && params.textoLibre.trim()) hoja.getRange(fila, idx.TEXTO_LIBRE + 1).setValue(params.textoLibre);
      if (huboConversionModo) hoja.getRange(fila, idx.PUNTOS + 1).setValue('[]'); // se limpia el otro modo al convertir
    }

    const filaDataDespues = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
    const snapshotDespues = {
      numeroActa: filaDataDespues[idx.NUMERO_ACTA],
      numeroActaAnterior: filaDataDespues[idx.NUMERO_ACTA_ANTERIOR],
      fechaReunion: filaDataDespues[idx.FECHA_REUNION] ? Utilities.formatDate(new Date(filaDataDespues[idx.FECHA_REUNION]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      horaInicio: filaDataDespues[idx.HORA_INICIO],
      horaFin: filaDataDespues[idx.HORA_FIN],
      presentes: filaDataDespues[idx.PRESENTES],
      modoContenido: filaDataDespues[idx.MODO_CONTENIDO],
      puntos: filaDataDespues[idx.PUNTOS],
      textoLibre: filaDataDespues[idx.TEXTO_LIBRE] || ''
    };

    logAuditoriaActa_('EDITAR_ASENTADA', params.idRegistro, filaDataDespues[idx.NUMERO_ACTA], params.motivo, estadoActual, { antes: snapshotAntes, despues: snapshotDespues });

    if (huboConversionModo) {
      logAuditoriaActa_('CONVERSION_MODO', params.idRegistro, filaDataDespues[idx.NUMERO_ACTA], params.motivo, estadoActual, { modoAnterior: modoActual, modoNuevo: modoNuevo });
    }

    let advertenciaReferencias = null;
    if (numeroNuevo !== null) {
      logAuditoriaActa_('CAMBIO_NUMERACION', params.idRegistro, numeroNuevo, params.motivo, estadoActual, { numeroAnterior: numeroActual, numeroNuevo: numeroNuevo });

      // Solo informativo -- detecta pero NO modifica ninguna otra fila
      let referencias = 0;
      for (let i = 1; i < datos.length; i++) {
        if (i === fila - 1) continue;
        if (String(datos[i][idx.NUMERO_ACTA_ANTERIOR]) === String(numeroActual)) referencias++;
      }
      if (referencias > 0) {
        advertenciaReferencias = 'Hay ' + referencias + ' acta(s) que todavía referencian al Acta N.º ' + numeroActual + ' como acta anterior. Revisalas si corresponde.';
      }
    }

    return { ok: true, advertencia: advertenciaReferencias };
  } finally {
    if (conseguido) lock.releaseLock();
  }
}

// MIGRACIÓN ÚNICA -- correr una sola vez para pasar del modelo viejo (ID_ACTA = número
// definitivo desde el nacimiento) al modelo nuevo (ID_REGISTRO técnico + NUMERO_ACTA diferido).
// Es idempotente: si detecta que ya migró (existe columna ID_REGISTRO con datos), no hace nada.
// MIGRACIÓN ADITIVA -- agrega MODO_CONTENIDO y TEXTO_LIBRE a ACTAS sin tocar contenido ni
// numeración de nada existente. Todo lo migrado (y todo lo previo) queda en ESTRUCTURADO,
// que es exactamente el comportamiento que ya tenían. Idempotente.
function migrarActasV3ModoContenido() {
  const hoja = getHojaActas();
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

  if (headers.indexOf('MODO_CONTENIDO') !== -1) {
    return { ok: true, mensaje: 'Ya está migrado (columna MODO_CONTENIDO ya existe). No se hizo nada.' };
  }

  const colBase = hoja.getLastColumn();
  hoja.getRange(1, colBase + 1).setValue('MODO_CONTENIDO');
  hoja.getRange(1, colBase + 2).setValue('TEXTO_LIBRE');
  hoja.getRange(1, colBase + 1, 1, 2).setFontWeight('bold').setBackground('#135457').setFontColor('#c4df57');

  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  let marcadas = 0;
  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][idx.ID_REGISTRO]) continue;
    hoja.getRange(i + 1, idx.MODO_CONTENIDO + 1).setValue('ESTRUCTURADO');
    hoja.getRange(i + 1, idx.TEXTO_LIBRE + 1).setValue('');
    marcadas++;
  }

  return { ok: true, mensaje: 'Columnas agregadas. ' + marcadas + ' acta(s) marcada(s) como ESTRUCTURADO (comportamiento sin cambios).' };
}

function migrarActasV2() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = getHojaActas();
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

  if (headers.indexOf('ID_REGISTRO') !== -1) {
    return { ok: true, mensaje: 'Ya está migrado (columna ID_REGISTRO ya existe). No se hizo nada.' };
  }

  getHojaAuditoriaActas(); // se asegura de que exista

  // Guarda snapshot de datos actuales (id numérico viejo) ANTES de tocar headers, para no perder referencia
  const datosViejos = hoja.getDataRange().getValues();
  const idxViejo = {};
  datosViejos[0].forEach((h, i) => idxViejo[h] = i);

  // Renombra headers existentes en el lugar (misma posición de columna, solo cambia el texto)
  hoja.getRange(1, idxViejo.ID_ACTA + 1).setValue('NUMERO_ACTA');
  hoja.getRange(1, idxViejo.ID_ACTA_ANTERIOR + 1).setValue('NUMERO_ACTA_ANTERIOR');

  // Agrega columnas nuevas al final
  const colBase = hoja.getLastColumn();
  hoja.getRange(1, colBase + 1).setValue('ID_REGISTRO');
  hoja.getRange(1, colBase + 2).setValue('ORIGEN');
  hoja.getRange(1, colBase + 3).setValue('FECHA_ASENTAMIENTO');
  hoja.getRange(1, colBase + 4).setValue('MOTIVO_ANULACION');
  hoja.getRange(1, colBase + 1, 1, 4).setFontWeight('bold').setBackground('#135457').setFontColor('#c4df57');

  // Migra BITACORA: renombra columna y prepara mapa numero-viejo -> idRegistro-nuevo (se completa abajo)
  const hojaBit = getHojaBitacora();
  const headersBit = hojaBit.getRange(1, 1, 1, hojaBit.getLastColumn()).getValues()[0];
  const idxBitViejo = {};
  headersBit.forEach((h, i) => idxBitViejo[h] = i);
  if (headersBit.indexOf('ID_ACTA_DESTINO') !== -1) {
    hojaBit.getRange(1, idxBitViejo.ID_ACTA_DESTINO + 1).setValue('ID_REGISTRO_DESTINO');
  }

  const mapaNumeroAIdRegistro = {}; // { '556': 'ACTA-MIG-556', ... }
  const numerosAPurgar = [564, 565, 566]; // pruebas de hoy, confirmadas para purgar
  const numeroAConservar = 563; // borrador real de hoy, se conserva

  // Re-lee datos ya con headers nuevos, en el mismo orden de filas
  const datosActuales = hoja.getDataRange().getValues();
  const idx = {};
  datosActuales[0].forEach((h, i) => idx[h] = i);

  const filasAPurgar = []; // se borran al final, de abajo hacia arriba

  for (let i = 1; i < datosActuales.length; i++) {
    const fila = datosActuales[i];
    const numeroViejo = Number(fila[idx.NUMERO_ACTA]); // toma temporalmente el valor viejo, ya que la columna se llama distinto pero el dato sigue ahí
    const estadoViejo = fila[idx.ESTADO];
    const filaSheet = i + 1;

    if (estadoViejo === 'FIRMADA') {
      const idRegistro = 'ACTA-MIG-' + numeroViejo;
      hoja.getRange(filaSheet, idx.ID_REGISTRO + 1).setValue(idRegistro);
      hoja.getRange(filaSheet, idx.ORIGEN + 1).setValue('MIGRADO');
      hoja.getRange(filaSheet, idx.ESTADO + 1).setValue('ASENTADA');
      mapaNumeroAIdRegistro[String(numeroViejo)] = idRegistro;
    } else if (estadoViejo === 'BORRADOR' && numeroViejo === numeroAConservar) {
      const idRegistro = 'ACTA-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-MIG563';
      hoja.getRange(filaSheet, idx.ID_REGISTRO + 1).setValue(idRegistro);
      hoja.getRange(filaSheet, idx.ORIGEN + 1).setValue('APP');
      hoja.getRange(filaSheet, idx.NUMERO_ACTA + 1).setValue(''); // pierde el número que tenía reservado bajo el modelo viejo
      mapaNumeroAIdRegistro[String(numeroViejo)] = idRegistro;
    } else if (estadoViejo === 'BORRADOR' && numerosAPurgar.indexOf(numeroViejo) !== -1) {
      filasAPurgar.push({ fila: filaSheet, numero: numeroViejo });
    }
  }

  // Migra BITACORA: reemplaza el número viejo por el idRegistro correspondiente en cada nota
  const datosBit = hojaBit.getDataRange().getValues();
  const idxBit = {};
  datosBit[0].forEach((h, i) => idxBit[h] = i);
  let notasMigradas = 0, notasLiberadas = 0;
  for (let i = 1; i < datosBit.length; i++) {
    const valorViejo = String(datosBit[i][idxBit.ID_REGISTRO_DESTINO] || '');
    if (!valorViejo) continue;
    if (mapaNumeroAIdRegistro[valorViejo]) {
      hojaBit.getRange(i + 1, idxBit.ID_REGISTRO_DESTINO + 1).setValue(mapaNumeroAIdRegistro[valorViejo]);
      notasMigradas++;
    } else if (numerosAPurgar.indexOf(Number(valorViejo)) !== -1) {
      hojaBit.getRange(i + 1, idxBit.PROCESADA + 1).setValue(false);
      hojaBit.getRange(i + 1, idxBit.ID_REGISTRO_DESTINO + 1).setValue('');
      notasLiberadas++;
    }
  }

  // Purga las actas de prueba, de abajo hacia arriba para no correr los índices de fila
  filasAPurgar.sort((a, b) => b.fila - a.fila);
  filasAPurgar.forEach(p => {
    logAuditoriaActa_('PURGA_MIGRACION', '', p.numero, 'Borrador de prueba, purgado en migración al modelo v2 (ID_REGISTRO + numeración diferida)', 'APP', null);
    hoja.deleteRow(p.fila);
  });

  return {
    ok: true,
    mensaje: 'Migración completa.',
    actasMigradasFirmadaAAsentada: Object.keys(mapaNumeroAIdRegistro).length - (mapaNumeroAIdRegistro[String(numeroAConservar)] ? 1 : 0),
    borradorConservado: mapaNumeroAIdRegistro[String(numeroAConservar)] || null,
    actasPurgadas: filasAPurgar.map(p => p.numero),
    notasBitacoraMigradas: notasMigradas,
    notasBitacoraLiberadas: notasLiberadas,
    proximoNumeroActaDisponible: obtenerUltimoNumeroActa() + 1
  };
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
  const estadoBalance = !balanceDelEjercicio ? 'NOINICIADO' : ((balanceDelEjercicio.estado === 'APROBADO_ASAMBLEA' || balanceDelEjercicio.estado === 'LEGALIZADO') ? 'COMPLETO' : (balanceDelEjercicio.estado === 'OBSERVADO' ? 'OBSERVADO' : 'PROGRESO'));
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

  const idEjercicioActivo = activo[idx.ID_EJERCICIO];

  // ── Prerrequisitos de cierre (Bloque 3) ──
  // BLOQUEA: sin esto, cerrar sería fabricar un cierre institucional que no ocurrió.
  const bloqueos = [];

  const hojaAsam = getHojaAsambleas();
  const datosAsam = hojaAsam.getDataRange().getValues();
  let filaAsambleaActiva = -1;
  for (let i = 1; i < datosAsam.length; i++) {
    if (datosAsam[i][1] === idEjercicioActivo) { filaAsambleaActiva = i; break; }
  }
  const asambleaCelebrada = filaAsambleaActiva !== -1 && datosAsam[filaAsambleaActiva][6] === 'CELEBRADA';
  if (!asambleaCelebrada) bloqueos.push('La Asamblea todavía no fue marcada como celebrada.');

  const resBalances = listarBalances();
  const balanceAprobado = resBalances.balances.some(b => b.idEjercicio === idEjercicioActivo && b.estado === 'APROBADO_ASAMBLEA');
  if (!balanceAprobado) bloqueos.push('El Balance todavía no tiene un registro con estado APROBADO_ASAMBLEA.');

  if (bloqueos.length > 0) {
    return { ok: false, error: 'No se puede cerrar el Ejercicio todavía.', bloqueos: bloqueos };
  }

  // ADVIERTE: se puede continuar si params.confirmarConAdvertencias === true.
  const advertencias = [];

  const resMemoria = listarDocumentos({ tipo: 'MEMORIA', idEjercicio: idEjercicioActivo });
  if (!resMemoria.documentos.length || resMemoria.documentos[0].estado !== 'APROBADA') {
    advertencias.push('La Memoria no está aprobada.');
  }

  const resInforme = listarDocumentos({ tipo: 'INFORME_REVISOR', idEjercicio: idEjercicioActivo });
  if (!resInforme.documentos.length) {
    advertencias.push('No se cargó el Informe del Revisor.');
  }

  const resActaAsam = listarDocumentos({ tipo: 'ACTA_ASAMBLEA', idEjercicio: idEjercicioActivo });
  if (!resActaAsam.documentos.length || resActaAsam.documentos[0].estado !== 'FIRMADA') {
    advertencias.push('El Acta de Asamblea no está firmada.');
  }

  const resAut = listarAutoridades();
  if (resAut.ok && resAut.vencenEsteAnio && resAut.vencenEsteAnio.length > 0) {
    const autoridadesRegistradas = filaAsambleaActiva !== -1 && datosAsam[filaAsambleaActiva][9] === 'SI';
    if (!autoridadesRegistradas) advertencias.push('Había autoridades por renovar este año y no fueron registradas.');
  }

  if (advertencias.length > 0 && !(params && params.confirmarConAdvertencias === 'true')) {
    return { ok: false, requiereConfirmacion: true, advertencias: advertencias, mensaje: 'Hay advertencias pendientes. Si querés cerrar de todas formas, volvé a llamar con confirmarConAdvertencias=true.' };
  }

  // ── A partir de acá, el cierre real (sin cambios respecto a la lógica anterior) ──

  // Cierra el actual
  hoja.getRange(filaActiva + 1, idx.ESTADO + 1).setValue('PRESENTADO');

  // Abre el siguiente: inicio = día después del cierre anterior, cierre = +1 año, límite asamblea = cierre + 4 meses (Art. 39)
  const cierreAnterior = new Date(activo[idx.FECHA_CIERRE]);
  const nuevoInicio = new Date(cierreAnterior);
  nuevoInicio.setDate(nuevoInicio.getDate() + 1);
  const nuevoCierre = new Date(nuevoInicio);
  nuevoCierre.setFullYear(nuevoCierre.getFullYear() + 1);
  nuevoCierre.setDate(nuevoCierre.getDate() - 1);
  const nuevoLimiteAsamblea = new Date(nuevoCierre);
  nuevoLimiteAsamblea.setMonth(nuevoLimiteAsamblea.getMonth() + 4);

  const nuevoNumero = Number(activo[idx.NUMERO]) + 1;
  const nuevoId = 'EJ-' + String(nuevoNumero).padStart(3, '0');

  hoja.appendRow([nuevoId, nuevoNumero, nuevoInicio, nuevoCierre, 'ABIERTO', nuevoLimiteAsamblea]);

  return { ok: true, ejercicioAnteriorCerrado: activo[idx.ID_EJERCICIO], nuevoEjercicio: nuevoId, nuevoNumero: nuevoNumero, advertenciasIgnoradas: advertencias };
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

// ── Bloque 2: fix estructural de ASAMBLEAS (fila de encabezado faltante) ──
// Correr en este orden exacto, uno por vez, verificando entre cada paso:
//   1) backupAsambleas()
//   2) insertarEncabezadoAsambleas()
//   3) verificar manualmente que las 2 filas sigan intactas debajo del encabezado
//   4) limpiarDuplicadosAsambleas() (ya existe, va a funcionar recién ahora)
//   5) verificar que quedó 1 sola fila, la del 24/10 con Orden del Día completo

function backupAsambleas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const original = ss.getSheetByName(HOJA_ASAMBLEAS);
  if (!original) return { ok: false, error: 'No existe la hoja ASAMBLEAS' };
  const nombreBackup = 'ASAMBLEAS_BACKUP_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const yaExiste = ss.getSheetByName(nombreBackup);
  if (yaExiste) return { ok: false, error: 'Ya existe un backup con ese nombre, esperá un segundo y reintentá' };
  const copia = original.copyTo(ss);
  copia.setName(nombreBackup);
  return { ok: true, mensaje: 'Backup creado como pestaña: ' + nombreBackup };
}

function backupDocumentos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const original = ss.getSheetByName(HOJA_DOCUMENTOS);
  if (!original) return { ok: false, error: 'No existe la hoja DOCUMENTOS' };
  const nombreBackup = 'DOCUMENTOS_BACKUP_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const yaExiste = ss.getSheetByName(nombreBackup);
  if (yaExiste) return { ok: false, error: 'Ya existe un backup con ese nombre, esperá un segundo y reintentá' };
  const copia = original.copyTo(ss);
  copia.setName(nombreBackup);
  return { ok: true, mensaje: 'Backup creado como pestaña: ' + nombreBackup };
}

// Solo diagnóstico -- NO borra ni modifica nada. Lista los duplicados
// tipo+idEjercicio+version encontrados en DOCUMENTOS para EJ-037.
function diagnosticoDuplicadosDocumentos() {
  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  const tiposRelevantes = ['ACTA_CD_CONVOCATORIA', 'EDICTO_DIARIO', 'EDICTO_BOLETIN', 'CIRCULAR'];
  const filas = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0]) continue;
    if (tiposRelevantes.indexOf(fila[2]) === -1) continue;
    filas.push({ fila: i + 1, idDocumento: fila[0], idEjercicio: fila[1], tipo: fila[2], version: fila[3], estado: fila[4], generadoPor: fila[6], fechaGeneracion: fila[7] });
  }
  return { ok: true, filas: filas };
}

// Saneamiento puntual (Bloque de cierre): borra los 4 documentos de Convocatoria
// generados el 11/08, ANTES de que existiera el fix de versionado -- confirmados como
// erróneos (mes en inglés, sin 2 autoridades que se cargaron después) y nunca aprobados.
// IDs hardcodeados a propósito, no lógica genérica -- acción quirúrgica, no automática.
function limpiarDocumentosConvocatoriaViejos() {
  const idsABorrar = [
    'DOC-20260811110603-ACTA_CD_CONVOCATORIA',
    'DOC-20260811110603-EDICTO_DIARIO',
    'DOC-20260811110604-EDICTO_BOLETIN',
    'DOC-20260811110605-CIRCULAR'
  ];
  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  const borrados = [];
  // De abajo hacia arriba, para no correr los índices al borrar
  for (let i = datos.length - 1; i >= 1; i--) {
    if (idsABorrar.indexOf(datos[i][0]) !== -1) {
      borrados.push({ idDocumento: datos[i][0], tipo: datos[i][2], version: datos[i][3] });
      hoja.deleteRow(i + 1);
    }
  }
  return { ok: true, borrados: borrados, cantidadBorrada: borrados.length, cantidadEsperada: idsABorrar.length };
}

function insertarEncabezadoAsambleas() {
  const hoja = getHojaAsambleas();
  const primeraFila = hoja.getRange(1, 1, 1, 1).getValue();
  if (primeraFila === 'ID_ASAMBLEA') {
    return { ok: false, error: 'Ya existe un encabezado (la fila 1 ya dice ID_ASAMBLEA). No se insertó nada, para no duplicarlo.' };
  }
  hoja.insertRowBefore(1);
  hoja.getRange(1, 1, 1, 10).setValues([[
    'ID_ASAMBLEA', 'ID_EJERCICIO', 'FECHA', 'HORA', 'LUGAR', 'ORDEN_DIA', 'ESTADO',
    'SOCIOS_PRESENTES', 'QUORUM_ALCANZADO', 'AUTORIDADES_REGISTRADAS'
  ]]);
  hoja.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#135457').setFontColor('#c4df57');
  return { ok: true, mensaje: 'Encabezado insertado como fila 1. Las filas de datos que había se corrieron una posición hacia abajo, sin modificarse.' };
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
        autoridadesRegistradas: datos[i][9] === 'SI'
      };
    }
  }
  return { ok: false, error: 'Todavía no hay fecha de Asamblea registrada para este Ejercicio' };
}

// FUENTE ÚNICA DE TEXTO para la Convocatoria. La página pública (convocatoria-publica.html)
// y el mail del Boletín (armarHTMLConvocatoria en el Código.gs del Boletín) tienen que verse
// distintos porque uno es una web y el otro un mail sin JS -- pero las FRASES tienen que ser
// las mismas siempre. Antes cada uno tenía su propia copia del texto y se desincronizaban
// (pasó el 14/08/2026). Ahora las dos piden esto acá -- para cambiar una palabra del texto de
// la Convocatoria, se edita UNA SOLA VEZ, en esta función.
function textosConvocatoria() {
  return {
    ok: true,
    saludo: 'Estimados socios:',
    intro: 'Por medio de la presente circular, conforme al artículo 41 del Estatuto Social, la Comisión Directiva convoca a la Asamblea General Ordinaria, que se realizará según los siguientes datos:',
    avisoQuorum: 'Conforme al Estatuto Social, si a la hora fijada no se reuniera la mayoría absoluta de los asociados con derecho a voto, la Asamblea se celebrará válidamente una hora después, con los socios presentes.',
    avisoCuotas: 'Solo pueden participar y votar los socios activos que se encuentren al día con el pago de sus cuotas sociales.',
    firmaLinea1: 'Comisión Directiva',
    firmaLinea2: 'Asociación Civil La Eugenia',
    sinOrdenDelDia: 'Todavía no se cargó el Orden del Día.'
  };
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
  // Art. 39 del Estatuto Social: "Las Asambleas Ordinarias deben convocarse dentro de los
  // cuatro (4) meses posteriores al cierre del ejercicio" -- confirmado contra el estatuto
  // vigente (no el que se venía usando antes, que decía 3 meses; era una versión desactualizada).
  const cierre = new Date(fechaCierreStr.split('/').reverse().join('-'));
  const limite = new Date(cierre);
  limite.setMonth(limite.getMonth() + 4);
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
        hojaAsam.getRange(i + 1, 10).setValue('SI'); // columna 10 = AUTORIDADES_REGISTRADAS
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
- Cierra con: "Siendo las [HORA] horas se da por finalizada la Asamblea, firmando la presente el Presidente de la Asamblea y los dos asociados designados para firmar el Acta.\\n\\nGarupá, [FECHA]."

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

// Arma el listado del Orden del Día como objetos con metadata, para que tanto la
// vista previa (Convocatoria) como la generación final de documentos usen exactamente
// la misma lógica y nunca queden desalineadas.
// tipo: 'FIJO' (obligatorio por Estatuto, se da siempre) | 'AUTOMATICO' (lo decide el
// sistema según datos reales, ej. si hay renovación de cargos) | 'MANUAL' (cargado a mano).
// Ninguno de FIJO/AUTOMATICO es editable desde la pantalla -- solo MANUAL.
// NOTA: el parámetro fueraDeTermino se conserva en la firma por compatibilidad con las llamadas
// existentes, pero ya no se usa -- el Estatuto (Art. 39) no exige que se justifique la demora
// en el Orden del Día, era una recomendación de prudencia, no una obligación. Se sacó a pedido
// explícito (18/08/2026).
function construirPuntosOrdenDelDia(ej, fueraDeTermino, vencenEsteAnio, puntosManuales) {
  const puntos = [];
  puntos.push({ tipo: 'FIJO', articulo: 'Art. 66', texto: 'Elección del Presidente de la Asamblea y de dos asociados para firmar el Acta.' });
  puntos.push({ tipo: 'FIJO', articulo: 'Art. 52.a', texto: 'Lectura y aprobación del Acta de la Asamblea anterior.' });
  puntos.push({ tipo: 'FIJO', articulo: 'Art. 52.b', texto: 'Consideración de la Memoria, Balance General, Inventario, Cuenta de Gastos y Recursos e Informe de la Comisión Revisora de Cuentas correspondientes al Ejercicio Económico N.° ' + ej.numero + ', finalizado el ' + ej.fechaCierre + '.' });
  puntos.push({ tipo: 'FIJO', articulo: 'Art. 16', texto: 'Aprobación del monto establecido por la Comisión Directiva para la cuota social, de conformidad con lo establecido en el artículo 16 del Estatuto Social. Definición del nuevo monto.' });
  if (vencenEsteAnio && vencenEsteAnio.length > 0) {
    puntos.push({ tipo: 'AUTOMATICO', articulo: 'Arts. 22, 35 y 63', texto: 'Renovación parcial de la Comisión Directiva y/o Comisión Revisora de Cuentas por finalización de mandato: ' + vencenEsteAnio.join(', ') + ', conforme a los artículos 22, 35 y 63 del Estatuto Social.' });
  }
  (puntosManuales || []).forEach(texto => {
    if (texto && texto.trim()) puntos.push({ tipo: 'MANUAL', articulo: '', texto: texto.trim() });
  });
  // Numeración final, en orden -- FIJO y AUTOMATICO ya vienen en el orden estatutario correcto,
  // los MANUAL siempre van al final (ver Art. 46: se incorporan antes de la convocatoria, no
  // reemplazan ni se mezclan con el temario obligatorio).
  puntos.forEach((p, i) => { p.numero = (i + 1) + 'º'; });
  return puntos;
}

// Vista previa de lectura -- no protegida, no escribe nada. Muestra el Orden del Día completo
// tal como quedaría, separando qué es fijo por Estatuto, qué decide el sistema solo, y qué
// fue cargado a mano, para que en la pantalla de Convocatoria se entienda qué se puede tocar.
function previsualizarOrdenDelDia() {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  const hojaAsam = getHojaAsambleas();
  const datosAsam = hojaAsam.getDataRange().getValues();
  let filaAsam = null;
  for (let i = 1; i < datosAsam.length; i++) {
    if (datosAsam[i][1] === ej.idEjercicio) { filaAsam = datosAsam[i]; break; }
  }
  if (!filaAsam) return { ok: false, error: 'Todavía no hay fecha de Asamblea cargada.' };

  const fechaAsamblea = new Date(filaAsam[2]);
  const fechaLimite = calcularFechaLimiteAsamblea(ej.fechaCierre);
  const fueraDeTermino = fechaAsamblea > fechaLimite;

  const autRes = listarAutoridades();
  const vencenEsteAnio = autRes.ok ? autRes.vencenEsteAnio : [];

  const puntosManuales = filaAsam[10] ? JSON.parse(filaAsam[10]) : [];

  const puntos = construirPuntosOrdenDelDia(ej, fueraDeTermino, vencenEsteAnio, puntosManuales);
  return { ok: true, puntos: puntos };
}

// Agrega una columna PUNTOS_MANUALES (11ª) a la hoja ASAMBLEAS si todavía no existe.
// Idempotente -- correrla de nuevo no rompe nada si ya está.
// Agrega ACTA_LEIDA_NUMERO a ACTAS -- distingue el hecho "qué acta se leyó en la reunión" de
// NUMERO_ACTA_ANTERIOR (correlatividad del Libro), que pueden diferir si entre la reunión y el
// asentamiento se incorpora otra acta (manual o de otro borrador). Idempotente.
// Además, de paso, etiqueta el Punto 1 de cada acta ESTRUCTURADO existente con
// tipo:'LECTURA_ACTA_ANTERIOR' cuando su texto coincide EXACTO con el patrón estándar -- esto es
// solo para la migración/reconocimiento inicial, nunca se usa como lógica permanente después.
// SOLO LECTURA -- no escribe nada, no requiere ninguna migración adicional. Lista el estado de
// etiquetado del Punto 1 (LECTURA_ACTA_ANTERIOR) de cada acta, para revisar manualmente los casos
// que la migración no pudo etiquetar con seguridad (patrón de texto no exacto).
function diagnosticarEtiquetadoActas() {
  const hoja = getHojaActas();
  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  const patronPunto1 = /acta\s*n\.?°?º?\s*(\d+)/i;

  const actas = [];
  const estructuradasSinPuntoTipificado = [];
  const estructuradasConMasDeUnoTipificado = [];
  const discrepanciasTextoVsActaLeida = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_REGISTRO]) continue;

    const modoContenido = fila[idx.MODO_CONTENIDO] || 'ESTRUCTURADO';
    const actaLeidaNumero = fila[idx.ACTA_LEIDA_NUMERO] || null;
    let puntos = [];
    try { puntos = fila[idx.PUNTOS] ? JSON.parse(fila[idx.PUNTOS]) : []; } catch (e) { puntos = null; }

    const puntosTipificados = Array.isArray(puntos) ? puntos.filter(p => p && p.tipo === 'LECTURA_ACTA_ANTERIOR').length : null;
    const puntoOrden1 = Array.isArray(puntos) ? puntos.find(p => p && p.orden === 1) : null;
    const textoOrden1 = puntoOrden1 ? puntoOrden1.texto : (modoContenido === 'TEXTO_LIBRE' ? '(TEXTO_LIBRE -- no aplica Punto 1)' : null);

    let numeroDetectadoEnTexto = null;
    let discrepancia = false;
    if (textoOrden1 && modoContenido === 'ESTRUCTURADO') {
      const match = String(textoOrden1).match(patronPunto1);
      if (match) {
        numeroDetectadoEnTexto = Number(match[1]);
        if (actaLeidaNumero !== null && numeroDetectadoEnTexto !== Number(actaLeidaNumero)) {
          discrepancia = true;
        }
      }
    }

    const registro = {
      idRegistro: fila[idx.ID_REGISTRO],
      numeroActa: fila[idx.NUMERO_ACTA] || null,
      estado: fila[idx.ESTADO],
      actaLeidaNumero: actaLeidaNumero,
      numeroActaAnterior: fila[idx.NUMERO_ACTA_ANTERIOR] || null,
      modoContenido: modoContenido,
      puntosTipificados: puntosTipificados,
      textoOrden1: textoOrden1,
      numeroDetectadoEnTexto: numeroDetectadoEnTexto,
      discrepanciaTextoVsActaLeida: discrepancia
    };
    actas.push(registro);

    if (modoContenido === 'ESTRUCTURADO' && puntosTipificados === 0) {
      estructuradasSinPuntoTipificado.push({ idRegistro: registro.idRegistro, numeroActa: registro.numeroActa, estado: registro.estado, textoOrden1: registro.textoOrden1 });
    }
    if (modoContenido === 'ESTRUCTURADO' && puntosTipificados > 1) {
      estructuradasConMasDeUnoTipificado.push({ idRegistro: registro.idRegistro, numeroActa: registro.numeroActa, estado: registro.estado });
    }
    if (discrepancia) {
      discrepanciasTextoVsActaLeida.push({ idRegistro: registro.idRegistro, numeroActa: registro.numeroActa, actaLeidaNumero: actaLeidaNumero, numeroDetectadoEnTexto: numeroDetectadoEnTexto, textoOrden1: registro.textoOrden1 });
    }
  }

  return {
    ok: true,
    actas: actas,
    resumen: {
      totalActas: actas.length,
      estructuradasSinPuntoTipificado: estructuradasSinPuntoTipificado,
      estructuradasConMasDeUnoTipificado: estructuradasConMasDeUnoTipificado,
      discrepanciasTextoVsActaLeida: discrepanciasTextoVsActaLeida
    }
  };
}
function migrarColumnaActaLeida() {
  const hoja = getHojaActas();
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  if (headers.indexOf('ACTA_LEIDA_NUMERO') !== -1) return { ok: true, mensaje: 'Ya existía.' };

  const colBase = hoja.getLastColumn();
  hoja.getRange(1, colBase + 1).setValue('ACTA_LEIDA_NUMERO').setFontWeight('bold').setBackground('#135457').setFontColor('#c4df57');

  const datos = hoja.getDataRange().getValues();
  const idx = {};
  datos[0].forEach((h, i) => idx[h] = i);

  const patronPunto1 = /^Se da lectura al acta N\.º (\d+)\. Se aprueba por unanimidad\.$/;
  let asentadasCompletadas = 0, puntosEtiquetados = 0;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[idx.ID_REGISTRO]) continue;
    const filaSheet = i + 1;
    const estado = fila[idx.ESTADO];

    // ACTA_LEIDA_NUMERO: solo se completa para ASENTADA/ANULADA, copiando NUMERO_ACTA_ANTERIOR
    // si existe. Los BORRADOR quedan vacíos a propósito -- no se inventa ningún valor.
    if ((estado === 'ASENTADA' || estado === 'ANULADA') && fila[idx.NUMERO_ACTA_ANTERIOR]) {
      hoja.getRange(filaSheet, colBase + 1).setValue(fila[idx.NUMERO_ACTA_ANTERIOR]);
      asentadasCompletadas++;
    }

    // Etiquetado del Punto 1 -- solo si coincide EXACTO con el patrón estándar, y solo para
    // actas ESTRUCTURADO. Si no coincide con seguridad, no se toca (queda sin tipo, para revisar
    // a mano antes de Asentar si hace falta).
    const modo = fila[idx.MODO_CONTENIDO] || 'ESTRUCTURADO';
    if (modo === 'ESTRUCTURADO' && fila[idx.PUNTOS]) {
      try {
        const puntos = JSON.parse(fila[idx.PUNTOS]);
        if (Array.isArray(puntos) && puntos.length > 0 && puntos[0].orden === 1 && !puntos[0].tipo) {
          const match = String(puntos[0].texto || '').trim().match(patronPunto1);
          if (match) {
            puntos[0].tipo = 'LECTURA_ACTA_ANTERIOR';
            hoja.getRange(filaSheet, idx.PUNTOS + 1).setValue(JSON.stringify(puntos));
            puntosEtiquetados++;
          }
        }
      } catch (e) { /* PUNTOS mal formado -- se ignora esa fila, no rompe la migración */ }
    }
  }

  return { ok: true, mensaje: 'Columna agregada. ' + asentadasCompletadas + ' acta(s) definitiva(s) completadas, ' + puntosEtiquetados + ' Punto 1 etiquetado(s) como LECTURA_ACTA_ANTERIOR.' };
}


function migrarColumnaPuntosManuales() {
  const hoja = getHojaAsambleas();
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  if (headers.indexOf('PUNTOS_MANUALES') !== -1) return { ok: true, mensaje: 'Ya existía.' };
  const nuevaCol = hoja.getLastColumn() + 1;
  hoja.getRange(1, nuevaCol).setValue('PUNTOS_MANUALES').setFontWeight('bold').setBackground('#135457').setFontColor('#c4df57');
  return { ok: true, mensaje: 'Columna agregada en la posición ' + nuevaCol + '.' };
}

// Guarda el listado completo de puntos manuales (reemplaza el anterior) para el Ejercicio activo.
function guardarPuntosManualesAsamblea(params) {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const hoja = getHojaAsambleas();
  const datos = hoja.getDataRange().getValues();
  const puntosManuales = params.puntosManuales ? JSON.parse(params.puntosManuales) : [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === ejercicioActivo.ejercicio.idEjercicio) {
      hoja.getRange(i + 1, 11).setValue(JSON.stringify(puntosManuales));
      return { ok: true };
    }
  }
  return { ok: false, error: 'No hay Asamblea registrada para este Ejercicio. Cargá la fecha primero en el Home.' };
}

// Arma los 4 textos (Acta CD, Edicto Diario, Edicto Boletín, Circular) a partir del Orden del
// Día recalculado en el momento. Compartida por generarConvocatoriaYDocumentos (crea versión
// nueva) y actualizarOrdenDelDiaEnDocumentos (sobrescribe en el lugar) -- así nunca hay dos
// copias de esta plantilla que se puedan desincronizar entre sí.
function armarTextosConvocatoria_(ej, params, numeroActaUsar) {
  const hojaAsam = getHojaAsambleas();
  const datosAsam = hojaAsam.getDataRange().getValues();
  let filaExistente = -1;
  for (let i = 1; i < datosAsam.length; i++) {
    if (datosAsam[i][1] === ej.idEjercicio) { filaExistente = i + 1; break; }
  }
  if (filaExistente === -1) {
    return { ok: false, error: 'Todavía no hay fecha de Asamblea cargada. Cargala primero en la píldora del Home.' };
  }
  const filaAsam = datosAsam[filaExistente - 1];
  const fechaAsamblea = new Date(filaAsam[2]);
  const horaAsamblea = formatearHoraSegura(filaAsam[3]);
  const lugarAsamblea = filaAsam[4];

  const fechaLimite = calcularFechaLimiteAsamblea(ej.fechaCierre);
  const fueraDeTermino = fechaAsamblea > fechaLimite;

  const autRes = listarAutoridades();
  const vencenEsteAnio = autRes.ok ? autRes.vencenEsteAnio : [];

  const ultimoNumero = obtenerUltimoNumeroActa();

  // Orden del Día -- objeto único, todos los documentos lo citan literal, nunca se retipea
  const puntosManuales = filaAsam[10] ? JSON.parse(filaAsam[10]) : [];
  const puntosObj = construirPuntosOrdenDelDia(ej, fueraDeTermino, vencenEsteAnio, puntosManuales);
  const puntos = puntosObj.map(p => p.numero + ' ' + p.texto);
  const ordenDelDiaTexto = puntos.join('\n');

  // Mantiene ASAMBLEAS.ORDEN_DIA sincronizado con lo último calculado, tanto al generar
  // como al actualizar -- la página pública y el mail del Boletín leen de acá.
  hojaAsam.getRange(filaExistente, 6).setValue(JSON.stringify(puntos));

  const MESES_ES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const fechaAsambleaFmt = Utilities.formatDate(fechaAsamblea, Session.getScriptTimeZone(), 'dd') + ' de ' + MESES_ES[fechaAsamblea.getMonth()] + ' de ' + Utilities.formatDate(fechaAsamblea, Session.getScriptTimeZone(), 'yyyy');

  // ---- 1. Acta de Comisión Directiva ----
  // El número solo se imprime cuando el acta está realmente ASENTADA en el módulo Actas.
  // Mientras esta pieza de Convocatoria no tenga vínculo real (no está implementado todavía),
  // nunca puede ser 'asentada' desde acá -- siempre es un texto provisorio, así que el
  // encabezado nunca debe mostrar un número como si fuera definitivo.
  let actaCD = 'ACTA N.º [A ASIGNAR AL ASENTAR]\n\n';
  actaCD += 'En la sede social del Club de Campo La Eugenia, siendo las ' + params.horaReunionCD + ' hs. del día ' + params.fechaReunionCD + ', se reúnen los siguientes miembros de Comisión Directiva: ' + params.presentesCD + '.\n\n';
  actaCD += 'Como primer punto del orden del día se da lectura al acta N.º ' + ultimoNumero + '. Se aprueba por unanimidad.\n\n';
  actaCD += 'Como segundo punto del orden del día, la Comisión Directiva resuelve convocar a Asamblea General Ordinaria para el día ' + fechaAsambleaFmt + ', a las ' + horaAsamblea + ' horas, en ' + lugarAsamblea + '.\n\n';
  actaCD += 'Como tercer punto del orden del día, se aprueba el siguiente Orden del Día para la Asamblea:\n\n' + ordenDelDiaTexto + '\n\n';
  actaCD += 'Siendo las ' + (params.horaFinReunionCD || '21:15') + ' horas se levanta la sesión.-';

  // ---- 2. Edicto diario local ----
  let edictoDiario = 'CLUB DE CAMPO "LA EUGENIA"\nConvocatoria a Asamblea General Ordinaria\n\n';
  edictoDiario += 'La Comisión Directiva convoca a los señores asociados a la Asamblea General Ordinaria, que se celebrará el día ' + fechaAsambleaFmt + ', a las ' + horaAsamblea + ' horas, en ' + lugarAsamblea + ', para tratar el siguiente:\n\nORDEN DEL DÍA\n\n' + ordenDelDiaTexto;
  edictoDiario += '\n\nConforme al artículo 44 del Estatuto Social, si a la hora fijada no se reuniera la mayoría absoluta de los asociados con derecho a voto, la Asamblea se celebrará válidamente una hora después.\n\nGarupá, Misiones.';

  // ---- 3. Edicto Boletín Oficial ----
  let edictoBoletin = edictoDiario + '\n\n[Publíquese 2 (dos) días en el Boletín Oficial de la Provincia de Misiones. El plazo de 15 días corridos se cuenta desde la última publicación hasta la Asamblea.]';

  // ---- 4. Circular a socios ----
  let circular = 'CIRCULAR\n\nSeñores Socios:\n\nPor medio de la presente circular, conforme al artículo 41 del Estatuto Social, la Comisión Directiva convoca a la Asamblea General Ordinaria, que se celebrará el día ' + fechaAsambleaFmt + ', a las ' + horaAsamblea + ' horas, en ' + lugarAsamblea + ', para tratar el siguiente:\n\nORDEN DEL DÍA\n\n' + ordenDelDiaTexto;
  circular += '\n\nRecordamos que, conforme a los artículos 13 y 47 del Estatuto Social, solo podrán participar y votar los socios activos que se encuentren al día con el pago de sus cuotas sociales.\n\nCOMISIÓN DIRECTIVA';

  return {
    ok: true,
    fueraDeTermino: fueraDeTermino,
    fechaLimite: fechaLimite,
    documentos: [
      { tipo: 'ACTA_CD_CONVOCATORIA', contenido: actaCD },
      { tipo: 'EDICTO_DIARIO', contenido: edictoDiario },
      { tipo: 'EDICTO_BOLETIN', contenido: edictoBoletin },
      { tipo: 'CIRCULAR', contenido: circular }
    ]
  };
}

function generarConvocatoriaYDocumentos(params) {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  const ultimoNumero = obtenerUltimoNumeroActa();
  const nuevoNumeroActa = ultimoNumero + 1;

  const armado = armarTextosConvocatoria_(ej, params, nuevoNumeroActa);
  if (!armado.ok) return armado;

  // Guarda los 4 documentos en la tabla única. La versión se calcula por tipo (no fija en 1)
  // para que, al volver a generar después de una aprobación, quede claro cuál es la vigente
  // -- listarDocumentos ordena por versión descendente para decidir "la última".
  const hoja = getHojaDocumentos();
  const datosDocsExistentes = hoja.getDataRange().getValues();
  function proximaVersionDoc_(tipo) {
    let v = 0;
    for (let i = 1; i < datosDocsExistentes.length; i++) {
      if (datosDocsExistentes[i][0] && datosDocsExistentes[i][2] === tipo && datosDocsExistentes[i][1] === ej.idEjercicio) {
        v = Math.max(v, Number(datosDocsExistentes[i][3]) || 0);
      }
    }
    return v + 1;
  }
  const resultado = [];
  armado.documentos.forEach(doc => {
    const version = proximaVersionDoc_(doc.tipo);
    const nuevoId = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + doc.tipo;
    hoja.appendRow([nuevoId, ej.idEjercicio, doc.tipo, version, 'BORRADOR', doc.contenido, 'IA', new Date()]);
    resultado.push({ tipo: doc.tipo, idDocumento: nuevoId, contenido: doc.contenido, estado: 'BORRADOR', version: version });
  });

  return { ok: true, documentos: resultado, numeroActa: nuevoNumeroActa };
}

// "Actualizar" en vez de "Generar de nuevo": sobrescribe el CONTENIDO de la versión vigente de
// cada uno de los 4 documentos con el Orden del Día recalculado, sin crear una versión nueva ni
// cambiar el número de Acta. Si alguno estaba APROBADO, vuelve a BORRADOR -- cambió el contenido,
// tiene que revisarse de nuevo antes de mandarse. Pensada para cuando solo cambió el Orden del
// Día (agregaste/sacaste un punto manual) y no hace falta repetir toda la carga de la reunión de CD.
function actualizarOrdenDelDiaEnDocumentos(params) {
  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  const idx = {}; // ID_DOCUMENTO, ID_EJERCICIO, TIPO, VERSION, ESTADO, CONTENIDO, ORIGEN, FECHA
  ['ID_DOCUMENTO','ID_EJERCICIO','TIPO','VERSION','ESTADO','CONTENIDO','ORIGEN','FECHA'].forEach((h, i) => idx[h] = i);

  const tipos = ['ACTA_CD_CONVOCATORIA', 'EDICTO_DIARIO', 'EDICTO_BOLETIN', 'CIRCULAR'];
  const filaVigentePorTipo = {};
  tipos.forEach(t => {
    let mejorFila = -1, mejorVersion = -1;
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][idx.TIPO] === t && datos[i][idx.ID_EJERCICIO] === ej.idEjercicio) {
        const v = Number(datos[i][idx.VERSION]) || 0;
        if (v > mejorVersion) { mejorVersion = v; mejorFila = i + 1; }
      }
    }
    filaVigentePorTipo[t] = mejorFila;
  });

  if (tipos.some(t => filaVigentePorTipo[t] === -1)) {
    return { ok: false, error: 'Todavía no generaste los 4 documentos. Usá "Generar los 4 documentos" primero -- "Actualizar" solo refresca lo que ya existe.' };
  }

  // El número de Acta CD no cambia por actualizar el Orden del Día -- se extrae del documento existente.
  const filaActaExistente = filaVigentePorTipo['ACTA_CD_CONVOCATORIA'];
  const contenidoActaExistente = String(hoja.getRange(filaActaExistente, idx.CONTENIDO + 1).getValue());
  const matchNumero = contenidoActaExistente.match(/^ACTA N\.º (\d+)/);
  const numeroActaUsar = matchNumero ? matchNumero[1] : obtenerUltimoNumeroActa();

  const armado = armarTextosConvocatoria_(ej, params, numeroActaUsar);
  if (!armado.ok) return armado;

  const resultado = [];
  armado.documentos.forEach(doc => {
    const fila = filaVigentePorTipo[doc.tipo];
    const estadoAnterior = hoja.getRange(fila, idx.ESTADO + 1).getValue();
    const estadoNuevo = estadoAnterior === 'APROBADO' ? 'BORRADOR' : estadoAnterior;
    hoja.getRange(fila, idx.CONTENIDO + 1).setValue(doc.contenido);
    hoja.getRange(fila, idx.ESTADO + 1).setValue(estadoNuevo);
    hoja.getRange(fila, idx.FECHA + 1).setValue(new Date());
    const idDocumento = hoja.getRange(fila, idx.ID_DOCUMENTO + 1).getValue();
    resultado.push({ tipo: doc.tipo, idDocumento: idDocumento, contenido: doc.contenido, estado: estadoNuevo, reabierto: estadoAnterior === 'APROBADO' });
  });

  return { ok: true, documentos: resultado, numeroActa: numeroActaUsar };
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

function eliminarDocumento(params) {
  if (!params.idDocumento) return { ok: false, error: 'ID requerido' };
  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === params.idDocumento) {
      hoja.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'No encontrado' };
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
