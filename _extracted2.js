const EXEC_URL_GI = 'https://script.google.com/macros/s/AKfycbyBxlDAteVncpf2hyQ-Nna6YZjMiZodntG9gZZ9aPXPei1As_yhqMUCnjnM9mHmJYYZJg/exec';
const EXEC_URL_PINES_CGV_CHECK = 'https://script.google.com/macros/s/AKfycbyt4PoL_IWHaXRlMNDy2cPtWmvj7f0-DLpSe5IHir54yPDNg8osPtX_0ezoopC4aGR_7A/exec';

let PIN_SESION = '';
let ACTA_ACTUAL = null;

document.addEventListener('DOMContentLoaded', function() {
  const savedPin = localStorage.getItem('cgv_pin');
  if (!savedPin) return;
  PIN_SESION = savedPin;
  jsonp(EXEC_URL_PINES_CGV_CHECK + '?action=validarPin&pin=' + encodeURIComponent(savedPin) + '&modulo=gestion-institucional', function(res) {
    if (res && res.ok && res.valido) {
      entrarApp();
    }
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
  cargarNotas();
  cargarActas();
}

function cerrarSesion() {
  localStorage.removeItem('cgv_pin');
  location.reload();
}

function cambiarTab(tab) {
  document.getElementById('tabBitacoraBtn').classList.toggle('active', tab === 'bitacora');
  document.getElementById('tabActasBtn').classList.toggle('active', tab === 'actas');
  document.getElementById('panelBitacora').classList.toggle('active', tab === 'bitacora');
  document.getElementById('panelActas').classList.toggle('active', tab === 'actas');
}

// ===== BITÁCORA =====
function agregarNota() {
  const texto = document.getElementById('notaTexto').value.trim();
  if (!texto) return;
  document.getElementById('loadingText').textContent = 'Guardando nota...';
  jsonp(EXEC_URL_GI + '?action=guardarNota&texto=' + encodeURIComponent(texto) + '&cargadoPor=' + encodeURIComponent(PIN_SESION), function(res) {
    document.getElementById('loadingText').textContent = '';
    if (res && res.ok) {
      document.getElementById('notaTexto').value = '';
      cargarNotas();
    } else {
      alert('Error: ' + (res && res.error ? res.error : 'desconocido'));
    }
  });
}

function cargarNotas() {
  jsonp(EXEC_URL_GI + '?action=listarNotasPendientes', function(res) {
    const cont = document.getElementById('listaNotas');
    cont.innerHTML = '';
    if (!res || !res.ok || res.notas.length === 0) {
      cont.innerHTML = '<div class="empty-msg">No hay notas pendientes.</div>';
      return;
    }
    res.notas.forEach(n => {
      const div = document.createElement('div');
      div.className = 'nota-card';
      div.innerHTML = '<div class="nota-meta">' + n.fechaCarga + '</div><div class="nota-texto">' + escapeHtml(n.texto) + '</div>';
      cont.appendChild(div);
    });
  });
}

function escapeHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

// ===== ACTAS =====
function toggleFormNuevaActa() {
  const f = document.getElementById('formNuevaActa');
  f.style.display = f.style.display === 'none' ? 'block' : 'none';
  if (!document.getElementById('fFechaReunion').value) {
    document.getElementById('fFechaReunion').value = new Date().toISOString().split('T')[0];
  }
}

function generarBorrador() {
  const fecha = document.getElementById('fFechaReunion').value;
  const hora = document.getElementById('fHoraInicio').value;
  const presentes = document.getElementById('fPresentes').value;
  if (!fecha || !presentes) { alert('Completá fecha y presentes'); return; }
  document.getElementById('loadingText').textContent = 'Generando borrador...';
  const params = 'action=generarBorradorActa&fechaReunion=' + encodeURIComponent(fecha) + '&horaInicio=' + encodeURIComponent(hora) + '&presentes=' + encodeURIComponent(presentes);
  jsonp(EXEC_URL_GI + '?' + params, function(res) {
    document.getElementById('loadingText').textContent = '';
    if (res && res.ok) {
      toggleFormNuevaActa();
      document.getElementById('fPresentes').value = '';
      cargarActas();
      cargarNotas();
      abrirEditor(res.idActa);
    } else {
      alert('Error: ' + (res && res.error ? res.error : 'desconocido'));
    }
  });
}

function cargarActas() {
  jsonp(EXEC_URL_GI + '?action=listarActas', function(res) {
    const cont = document.getElementById('listaActas');
    cont.innerHTML = '';
    if (!res || !res.ok || res.actas.length === 0) {
      cont.innerHTML = '<div class="empty-msg">Todavía no hay actas generadas en el sistema.</div>';
      return;
    }
    res.actas.forEach(a => {
      const div = document.createElement('div');
      div.className = 'acta-card';
      div.onclick = () => abrirEditor(a.idActa);
      div.innerHTML = '<span class="acta-num">Acta N.º ' + a.idActa + ' — ' + a.fechaReunion + '</span><span class="estado-badge estado-' + a.estado + '">' + a.estado + '</span>';
      cont.appendChild(div);
    });
  });
}

function abrirEditor(idActa) {
  jsonp(EXEC_URL_GI + '?action=obtenerActa&idActa=' + idActa, function(res) {
    if (!res || !res.ok) { alert('No se pudo abrir el acta'); return; }
    ACTA_ACTUAL = res.acta;
    document.getElementById('editorActa').style.display = 'block';
    document.getElementById('editorActaTitulo').textContent = 'Editando Acta N.º ' + ACTA_ACTUAL.idActa + ' (' + ACTA_ACTUAL.estado + ')';
    renderPuntos();
    document.getElementById('textoFinalBox').style.display = 'none';
    document.getElementById('editorActa').scrollIntoView({ behavior: 'smooth' });
  });
}

function renderPuntos() {
  const cont = document.getElementById('puntosContainer');
  cont.innerHTML = '';
  ACTA_ACTUAL.puntos.forEach((p, i) => {
    const div = document.createElement('div');
    div.className = 'punto-item';
    div.innerHTML = '<div class="punto-label">Punto ' + (i + 1) + '</div><textarea data-idx="' + i + '" oninput="actualizarPuntoLocal(' + i + ', this.value)">' + escapeHtml(p.texto) + '</textarea>';
    cont.appendChild(div);
  });
}

function actualizarPuntoLocal(idx, valor) {
  ACTA_ACTUAL.puntos[idx].texto = valor;
}

function agregarPuntoVacio() {
  ACTA_ACTUAL.puntos.push({ orden: ACTA_ACTUAL.puntos.length + 1, texto: '' });
  renderPuntos();
}

function guardarCambiosActa() {
  const horaFin = document.getElementById('eHoraFin').value;
  document.getElementById('loadingText').textContent = 'Guardando...';
  const params = 'action=actualizarActa&idActa=' + ACTA_ACTUAL.idActa + '&puntos=' + encodeURIComponent(JSON.stringify(ACTA_ACTUAL.puntos)) + '&horaFin=' + encodeURIComponent(horaFin);
  jsonp(EXEC_URL_GI + '?' + params, function(res) {
    document.getElementById('loadingText').textContent = '';
    if (res && res.ok) {
      cargarActas();
    } else {
      alert('Error: ' + (res && res.error ? res.error : 'desconocido'));
    }
  });
}

function generarTextoFinal() {
  const ordinales = ['primer','segundo','tercer','cuarto','quinto','sexto','séptimo','octavo','noveno','décimo'];
  const fechaTxt = ACTA_ACTUAL.fechaReunion;
  const horaFin = document.getElementById('eHoraFin').value || '21:15';

  let texto = 'ACTA N.º ' + ACTA_ACTUAL.idActa + ':\n\n';
  texto += 'En la sede social del Club de Campo La Eugenia, siendo las ' + ACTA_ACTUAL.horaInicio + ' hs. del día ' + fechaTxt + ', se reúnen los siguientes miembros de Comisión Directiva: ' + ACTA_ACTUAL.presentes + '.\n\n';

  ACTA_ACTUAL.puntos.forEach((p, i) => {
    const ord = ordinales[i] || (i + 1) + 'º';
    if (i === 0) {
      texto += 'Como ' + ord + ' punto del orden del día, ' + p.texto.charAt(0).toLowerCase() + p.texto.slice(1) + '\n\n';
    } else {
      texto += 'Como ' + ord + ' punto del orden del día, ' + p.texto + '\n\n';
    }
  });

  texto += 'Siendo las ' + horaFin + ' horas se levanta la sesión.-';

  document.getElementById('textoFinalBox').textContent = texto;
  document.getElementById('textoFinalBox').style.display = 'block';
}

function cerrarEditor() {
  document.getElementById('editorActa').style.display = 'none';
  ACTA_ACTUAL = null;
}