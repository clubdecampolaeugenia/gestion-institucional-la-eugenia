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

function doGet(e) {
  const action = e.parameter.action;
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
