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
const ACCIONES_PROTEGIDAS = ['guardarAutoridad', 'guardarNota', 'generarBorradorActa', 'actualizarActa', 'procesarBalance', 'actualizarEstadoBalance'];

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
      resultado = { ok: true, version: 'v8-thinking-disabled-16k-09ago-1140' };
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
    puntos.push({ orden: puntos.length + 1, texto: fila[idxBit.TEXTO] });
    idsNotasUsadas.push({ row: i + 1, id: fila[idxBit.ID_NOTA] });
  }

  hojaActas.appendRow([
    nuevoNumero,
    params.idEjercicio || '',
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

  hoja.appendRow([
    nuevoId,
    '',
    urls,
    version,
    estadoInicial,
    JSON.stringify(analisis),
    JSON.stringify(analisis.observaciones)
  ]);

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
    balances.push({ idBalance: fila[idx.ID_BALANCE], version: fila[idx.VERSION], estado: fila[idx.ESTADO] });
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
