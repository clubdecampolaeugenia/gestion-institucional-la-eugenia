const EXEC_URL_PINES_CGV_CHECK = 'https://script.google.com/macros/s/AKfycbyt4PoL_IWHaXRlMNDy2cPtWmvj7f0-DLpSe5IHir54yPDNg8osPtX_0ezoopC4aGR_7A/exec';
const EXEC_URL_GI = 'https://script.google.com/macros/s/AKfycbyBxlDAteVncpf2hyQ-Nna6YZjMiZodntG9gZZ9aPXPei1As_yhqMUCnjnM9mHmJYYZJg/exec';

let PIN_SESION = '';

document.addEventListener('DOMContentLoaded', function() {
  // Si ya hay un PIN guardado de una sesión anterior en esta misma app, entra directo
  // sin volver a validar contra el servidor -- el PIN solo se procesa al entrar por primera vez.
  const savedPin = localStorage.getItem('cgv_pin');
  if (!savedPin) return;
  PIN_SESION = savedPin;
  entrarApp();
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
  cargarBalances();
}

function cerrarSesion() {
  localStorage.removeItem('cgv_pin');
  location.reload();
}

function procesarBalance() {
  const btn = document.getElementById('btnProcesar');
  btn.disabled = true;
  document.getElementById('loadingText').innerHTML = '<span class="spinner"></span> Leyendo el PDF y consultando a la IA... puede tardar hasta un minuto.';
  document.getElementById('resultadoContainer').innerHTML = '';

  jsonp(EXEC_URL_GI + '?action=procesarBalance&pin=' + encodeURIComponent(PIN_SESION), function(res) {
    btn.disabled = false;
    document.getElementById('loadingText').textContent = '';
    if (!res || !res.ok) {
      document.getElementById('resultadoContainer').innerHTML = '<div class="obs-item obs-CRITICO"><span class="obs-tipo">Error</span>' + escapeHtml(res && res.error ? res.error : 'desconocido') + '</div>';
      return;
    }
    mostrarResultado(res.analisis, res.estado);
    cargarBalances();
  });
}

function escapeHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function mostrarResultado(analisis, estado) {
  const cont = document.getElementById('resultadoContainer');
  let html = '<div class="resultado-box">';
  html += '<div class="resultado-titulo">Resultado del análisis</div>';

  if (analisis.documentosRecibidos && analisis.documentosRecibidos.length) {
    html += '<div class="dato-linea" style="margin-bottom:8px;"><strong>Documentos recibidos:</strong></div>';
    analisis.documentosRecibidos.forEach(d => {
      const esDefinitivo = d.tipo === 'ESTADOS_CONTABLES_DEFINITIVOS';
      const chipColor = esDefinitivo ? 'var(--verdecl)' : 'var(--naranja)';
      const chipLabel = esDefinitivo ? 'Estados Contables definitivos' : 'Hoja de trabajo interna';
      html += '<div style="display:inline-block; background:' + chipColor + '; border-radius:12px; padding:4px 10px; font-size:11px; font-weight:900; color:var(--petroleo); margin:0 6px 6px 0;">' + chipLabel + '</div>';
    });
  }

  html += '<div class="dato-linea"><strong>Entidad:</strong> ' + escapeHtml(analisis.denominacionEncontrada || '—') + '</div>';
  html += '<div class="dato-linea"><strong>Período:</strong> ' + escapeHtml(analisis.periodoInicio || '—') + ' a ' + escapeHtml(analisis.periodoCierre || '—') + '</div>';
  html += '<div class="dato-linea"><strong>Activo:</strong> $' + escapeHtml(String(analisis.activoTotal || '—')) + ' · <strong>Pasivo:</strong> $' + escapeHtml(String(analisis.pasivoTotal || '—')) + ' · <strong>PN:</strong> $' + escapeHtml(String(analisis.patrimonioNeto || '—')) + '</div>';
  html += '<div class="dato-linea"><strong>Superávit del ejercicio:</strong> $' + escapeHtml(String(analisis.superavitEjercicio || '—')) + '</div>';
  html += '<div class="dato-linea"><strong>Cuadra Activo = Pasivo + PN:</strong> ' + (analisis.cuadraActivoPasivoPN ? '✅ Sí' : '❌ No') + '</div>';
  html += '</div>';

  (analisis.observaciones || []).forEach(o => {
    html += '<div class="obs-item obs-' + o.tipo + '"><span class="obs-tipo">' + o.tipo.replace('_', ' ') + '</span>' + escapeHtml(o.texto) + '</div>';
  });

  cont.innerHTML = html;
}

function cargarBalances() {
  jsonp(EXEC_URL_GI + '?action=listarBalances', function(res) {
    const cont = document.getElementById('listaBalances');
    cont.innerHTML = '';
    if (!res || !res.ok || res.balances.length === 0) {
      cont.innerHTML = '<div class="empty-msg">Todavía no se procesó ningún balance.</div>';
      return;
    }
    res.balances.forEach(b => {
      const div = document.createElement('div');
      div.className = 'dato-linea';
      div.style.cssText = 'background:var(--blanco); border-radius:10px; padding:8px 12px; margin-bottom:6px; cursor:pointer;';
      div.textContent = 'Versión ' + b.version + ' — ' + b.estado;
      div.onclick = () => verDetalle(b.idBalance);
      cont.appendChild(div);
    });
  });
}

function verDetalle(idBalance) {
  jsonp(EXEC_URL_GI + '?action=obtenerBalance&idBalance=' + idBalance, function(res) {
    if (res && res.ok) {
      mostrarResultado(res.balance.datosExtraidos, res.balance.estado);
      document.getElementById('resultadoContainer').scrollIntoView({ behavior: 'smooth' });
    }
  });
}