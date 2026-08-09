const EXEC_URL_PINES_CGV_CHECK = 'https://script.google.com/macros/s/AKfycbyt4PoL_IWHaXRlMNDy2cPtWmvj7f0-DLpSe5IHir54yPDNg8osPtX_0ezoopC4aGR_7A/exec';
const EXEC_URL_GI = 'https://script.google.com/macros/s/AKfycbyBxlDAteVncpf2hyQ-Nna6YZjMiZodntG9gZZ9aPXPei1As_yhqMUCnjnM9mHmJYYZJg/exec';

let PIN_SESION = '';

// Hitos institucionales del ciclo anual, en orden.
// estado: COMPLETO / PROGRESO / OBSERVADO / PENDIENTE / NOINICIADO
// Los marcados "manual: true" reflejan el estado real conocido hoy; pasarán a calcularse solos
// cuando esos módulos (Balance, Memoria, Asamblea, Padrón) tengan su propio backend.
const HITOS = [
  { nombre: 'Ejercicio cerrado', detalle: '30/04/2026 — Ejercicio N.º 37', estado: 'COMPLETO', link: null, manual: true },
  { nombre: 'Balance', detalle: 'Recibido, con errores detectados (entidad y fechas)', estado: 'OBSERVADO', link: 'balance.html', manual: false },
  { nombre: 'Informe Comisión Revisora', detalle: 'Falta informe formal firmado (Art. 36)', estado: 'PENDIENTE', link: null, manual: true },
  { nombre: 'Memoria', detalle: 'Borrador redactado, falta aprobación de CD', estado: 'PROGRESO', link: null, manual: true },
  { nombre: 'Acta CD — aprobación y convocatoria', detalle: 'Borrador generado, pendiente de contrastar y firmar', estado: 'PROGRESO', link: null, manual: true },
  { nombre: 'Autoridades', detalle: 'cargando...', estado: 'PROGRESO', link: 'autoridades.html', manual: false },
  { nombre: 'Convocatoria y edictos', detalle: 'Borradores generados (diario + Boletín Oficial)', estado: 'PROGRESO', link: null, manual: true },
  { nombre: 'Padrón de socios', detalle: 'No iniciado — se arma leyendo Base Madre', estado: 'NOINICIADO', link: null, manual: true },
  { nombre: 'Bitácora y Actas', detalle: 'cargando...', estado: 'PROGRESO', link: 'actas.html', manual: false },
  { nombre: 'Asamblea General Ordinaria', detalle: 'Convocada para 24/10/2026, 16 hs', estado: 'PROGRESO', link: null, manual: true },
];

const ESTADO_LABEL = { COMPLETO: 'Completo', PROGRESO: 'En progreso', OBSERVADO: 'Observado', PENDIENTE: 'Pendiente', NOINICIADO: 'No iniciado' };
const ESTADO_PESO = { COMPLETO: 1, PROGRESO: 0.5, OBSERVADO: 0.25, PENDIENTE: 0, NOINICIADO: 0 };

document.addEventListener('DOMContentLoaded', function() {
  const savedPin = localStorage.getItem('cgv_pin');
  if (!savedPin) return;
  PIN_SESION = savedPin;
  jsonp(EXEC_URL_PINES_CGV_CHECK + '?action=validarPin&pin=' + encodeURIComponent(savedPin) + '&modulo=gestion-institucional', function(res) {
    if (res && res.ok && res.valido) entrarApp();
  });
});

function jsonp(url, cb) {
  const cbName = 'cb_' + Date.now() + Math.floor(Math.random()*100000);
  window[cbName] = function(res) { delete window[cbName]; document.body.removeChild(script); cb(res); };
  const script = document.createElement('script');
  script.src = url + (url.includes('?') ? '&' : '?') + 'callback=' + cbName;
  document.body.appendChild(script);
}

function verificarPin() {
  const pin = document.getElementById('pinInput').value;
  jsonp(EXEC_URL_PINES_CGV_CHECK + '?action=validarPin&pin=' + encodeURIComponent(pin) + '&modulo=gestion-institucional', function(res) {
    if (res && res.ok && res.valido) {
      PIN_SESION = pin;
      localStorage.setItem('cgv_pin', pin);
      entrarApp();
    } else {
      const input = document.getElementById('pinInput');
      input.classList.add('error');
      document.getElementById('pinError').textContent = 'PIN incorrecto';
      setTimeout(() => input.classList.remove('error'), 400);
    }
  });
}

function entrarApp() {
  document.getElementById('pinScreen').style.display = 'none';
  document.getElementById('mainScreen').style.display = 'block';
  document.getElementById('mainScreen').classList.add('active');
  renderHitos();
  cargarEstadoAutoridades();
  cargarEstadoActas();
  cargarEstadoBalance();
}

function cerrarSesion() {
  localStorage.removeItem('cgv_pin');
  location.reload();
}

function renderHitos() {
  const cont = document.getElementById('listaHitos');
  cont.innerHTML = '';
  HITOS.forEach((h, i) => {
    const el = document.createElement(h.link ? 'a' : 'div');
    el.className = 'hito-card' + (h.link ? ' clickable' : '');
    if (h.link) el.href = h.link;
    el.innerHTML =
      '<div class="hito-num">' + (i + 1) + '</div>' +
      '<div class="hito-info"><div class="hito-nombre">' + h.nombre + '</div><div class="hito-detalle" id="detalle-' + i + '">' + h.detalle + '</div></div>' +
      '<span class="hito-estado estado-' + h.estado.toLowerCase().replace('í','i') + '" id="estado-' + i + '">' + ESTADO_LABEL[h.estado] + '</span>';
    cont.appendChild(el);
  });
  actualizarProgreso();
}

function actualizarProgreso() {
  const total = HITOS.length;
  const suma = HITOS.reduce((acc, h) => acc + ESTADO_PESO[h.estado], 0);
  const pct = Math.round((suma / total) * 100);
  document.getElementById('progresoFill').style.width = pct + '%';
  document.getElementById('progresoPct').textContent = pct + '%';
}

function actualizarHito(nombre, estado, detalle) {
  const idx = HITOS.findIndex(h => h.nombre === nombre);
  if (idx === -1) return;
  HITOS[idx].estado = estado;
  HITOS[idx].detalle = detalle;
  document.getElementById('detalle-' + idx).textContent = detalle;
  const badge = document.getElementById('estado-' + idx);
  badge.className = 'hito-estado estado-' + estado.toLowerCase().replace('í','i');
  badge.textContent = ESTADO_LABEL[estado];
  actualizarProgreso();
}

function cargarEstadoAutoridades() {
  jsonp(EXEC_URL_GI + '?action=listarAutoridades', function(res) {
    if (!res || !res.ok) {
      actualizarHito('Autoridades', 'PENDIENTE', 'No se pudo consultar');
      return;
    }
    const vacantes = 12 - res.autoridades.length; // 11 cargos CD + al menos considerar Revisora/Arquitectura aparte
    const vencen = res.vencenEsteAnio.length;
    if (vacantes > 0) {
      actualizarHito('Autoridades', 'PROGRESO', vacantes + ' cargo(s) vacante(s) · ' + vencen + ' renuevan este año');
    } else {
      actualizarHito('Autoridades', 'COMPLETO', 'Todos los cargos cubiertos · ' + vencen + ' renuevan este año');
    }
  });
}

function cargarEstadoActas() {
  jsonp(EXEC_URL_GI + '?action=listarActas', function(res) {
    if (!res || !res.ok) {
      actualizarHito('Bitácora y Actas', 'NOINICIADO', 'Sin actas cargadas todavía');
      return;
    }
    const total = res.actas.length;
    const borradores = res.actas.filter(a => a.estado === 'BORRADOR').length;
    if (total === 0) {
      actualizarHito('Bitácora y Actas', 'NOINICIADO', 'Sin actas cargadas todavía');
    } else if (borradores > 0) {
      actualizarHito('Bitácora y Actas', 'PROGRESO', total + ' acta(s) — ' + borradores + ' en borrador');
    } else {
      actualizarHito('Bitácora y Actas', 'COMPLETO', total + ' acta(s), todas cerradas');
    }
  });
}

function cargarEstadoBalance() {
  jsonp(EXEC_URL_GI + '?action=listarBalances', function(res) {
    if (!res || !res.ok || res.balances.length === 0) {
      actualizarHito('Balance', 'NOINICIADO', 'Todavía no se procesó ningún balance');
      return;
    }
    const ultimo = res.balances[0];
    if (ultimo.estado === 'OBSERVADO') {
      actualizarHito('Balance', 'OBSERVADO', 'Versión ' + ultimo.version + ' — con observaciones de la IA');
    } else if (ultimo.estado === 'APROBADO_ASAMBLEA') {
      actualizarHito('Balance', 'COMPLETO', 'Versión ' + ultimo.version + ' — aprobado por Asamblea');
    } else {
      actualizarHito('Balance', 'PROGRESO', 'Versión ' + ultimo.version + ' — ' + ultimo.estado);
    }
  });
}