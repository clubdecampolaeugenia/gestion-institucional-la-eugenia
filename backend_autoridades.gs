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
const ACCIONES_PROTEGIDAS = ['guardarAutoridad', 'guardarNota', 'generarBorradorActa', 'actualizarActa', 'procesarBalance', 'actualizarEstadoBalance', 'cerrarYAbrirNuevoEjercicio', 'actualizarObservacionBalance', 'guardarNovedad', 'actualizarNovedad', 'generarBorradorMemoria', 'actualizarDocumento', 'guardarNovedadesSeleccionadas', 'extraerNovedadesDeChat', 'eliminarNovedad'];

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
      resultado = { ok: true, version: 'v15-notas-futuras-blindadas-09ago-2000' };
    } else if (action === 'obtenerEjercicioActivo') {
      resultado = obtenerEjercicioActivo();
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
    } else if (action === 'listarDocumentos') {
      resultado = listarDocumentos(e.parameter);
    } else if (action === 'actualizarDocumento') {
      resultado = actualizarDocumento(e.parameter);
    } else if (action === 'extraerNovedadesDeChat') {
      resultado = extraerNovedadesDeChat(e.parameter);
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
          horaInicio: fila[idx.HORA_INICIO],
          horaFin: fila[idx.HORA_FIN],
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

  // Marca las notas usadas como procesadas
  idsNotasUsadas.forEach(n => {
    hojaBitacora.getRange(n.row, idxBit.PROCESADA + 1).setValue(true);
    hojaBitacora.getRange(n.row, idxBit.ID_ACTA_DESTINO + 1).setValue(nuevoNumero);
  });

  return { ok: true, idActa: nuevoNumero, puntos: puntos };
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
      if (params.horaFin) hoja.getRange(row, idx.HORA_FIN + 1).setValue(params.horaFin);
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
function fechaEnRangoEjercicio(fechaStr, ejercicio) {
  const fecha = new Date(fechaStr);
  const inicio = new Date(ejercicio.fechaInicio.split('/').reverse().join('-'));
  const cierre = new Date(ejercicio.fechaCierre.split('/').reverse().join('-'));
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
    new Date(params.fecha),
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
      if (params.considerarMemoria) hoja.getRange(i + 1, idx.CONSIDERAR_MEMORIA + 1).setValue(params.considerarMemoria);
      if (params.titulo) hoja.getRange(i + 1, idx.TITULO + 1).setValue(params.titulo);
      if (params.descripcion) hoja.getRange(i + 1, idx.DESCRIPCION + 1).setValue(params.descripcion);
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

  return { ok: true, candidatos: candidatos };
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
      nuevoId, idEjercicioActual, new Date(n.fecha.split('/').reverse().join('-')),
      n.titulo, n.descripcion, n.categoria, n.montoSiCorresponde || '',
      'CHAT_IMPORTADO', '', considerarMemoria, ''
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
- Arranca siempre EXACTAMENTE así: "Señores Socios:\\n\\nEn cumplimiento de las disposiciones legales y estatutarias, ponemos a vuestra consideración la presente memoria, junto con los estados contables, el informe del auditor y el informe del revisor de cuentas correspondientes al ejercicio finalizado el [FECHA DE CIERRE]."
- El cuerpo es PROSA CORRIDA, en párrafos, SIN subtítulos ni encabezados de sección. Nunca uses títulos tipo "Seguridad:" o "Mantenimiento:" — los temas se mezclan en párrafos fluidos, ordenados por relevancia.
- Tono: primera persona plural ("hemos", "continuamos", "nos comprometemos"), formal pero cálido, nunca frío ni corporativo.
- Los temas se agrupan naturalmente en el texto (infraestructura, seguridad, EMSA/alumbrado, mantenimiento de espacios verdes, temporada de pileta, cantina, actividades sociales/deportivas, gestiones ante organismos como Municipalidad/EMSA/SAMSA/Personas Jurídicas), pero SIN anunciarlos como secciones.
- Los temas de "Personal" (altas, bajas, retribuciones, conflictos laborales) NUNCA se mencionan en la Memoria, aunque estén en las Novedades — no son parte de este documento.
- Cierra siempre con el desglose de socios que te paso como dato, con esta fórmula: "A la fecha de la presente memoria el número de socios Activos asciende a [N], de los cuales [N] abonan al día, mientras que [N] abona con un atraso de entre uno y cinco meses. Los socios No Activos, Morosos, son [N]."
- Después, SIEMPRE el agradecimiento: "La Comisión Directiva agradece profundamente a todos quienes colaboran con la gestión institucional y con el crecimiento sostenido del club, en especial al personal, cuyo compromiso y dedicación resultan fundamentales para el logro de los objetivos."
- Cierra con: "Garupá, [FECHA].\\n\\nCOMISIÓN DIRECTIVA"

Te paso las Novedades del ejercicio (ya filtradas a las que corresponde incluir) y los datos del Balance. Redactá el cuerpo integrando la información de forma natural y fluida, sin inventar nada que no esté en los datos que te doy. Si falta información importante para algún párrafo, simplemente no lo incluyas -- no inventes cifras ni hechos.

Respondé ÚNICAMENTE con el texto completo de la Memoria, sin explicaciones adicionales, sin markdown.`;

function generarBorradorMemoria(params) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'Falta configurar ANTHROPIC_API_KEY en Script Properties' };

  const ejercicioActivo = obtenerEjercicioActivo();
  if (!ejercicioActivo.ok) return { ok: false, error: 'No hay Ejercicio activo' };
  const ej = ejercicioActivo.ejercicio;

  // Junta las novedades SI/EVALUAR del ejercicio activo
  const novedadesRes = listarNovedades({ idEjercicio: ej.idEjercicio });
  const novedadesRelevantes = (novedadesRes.novedades || []).filter(n => n.considerarMemoria === 'SI' || n.considerarMemoria === 'EVALUAR');

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
    'NOVEDADES DEL EJERCICIO:\n' + novedadesRelevantes.map(n => '- [' + n.categoria + '] ' + n.titulo + ': ' + n.descripcion + (n.monto ? ' (monto: $' + n.monto + ')' : '')).join('\n');

  const payload = {
    model: 'claude-sonnet-5',
    max_tokens: 8000,
    thinking: { type: 'disabled' },
    system: PROMPT_MEMORIA,
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

  const hoja = getHojaDocumentos();
  const datos = hoja.getDataRange().getValues();
  let version = 1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] && datos[i][2] === 'MEMORIA' && datos[i][1] === ej.idEjercicio) version++;
  }

  const nuevoId = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  hoja.appendRow([nuevoId, ej.idEjercicio, 'MEMORIA', version, 'BORRADOR', bloqueTexto.text, 'IA', new Date()]);

  return { ok: true, idDocumento: nuevoId, contenido: bloqueTexto.text, novedadesUsadas: novedadesRelevantes.length };
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
