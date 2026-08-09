// ⚠️ Completar con la URL /exec del backend de Gestión Institucional una vez desplegado
const EXEC_URL_GI = 'PENDIENTE_DE_DEPLOY';
const EXEC_URL_PINES_CGV_CHECK = 'https://script.google.com/macros/s/AKfycbyt4PoL_IWHaXRlMNDy2cPtWmvj7f0-DLpSe5IHir54yPDNg8osPtX_0ezoopC4aGR_7A/exec';

const CARGOS_POR_ORGANO = {
  COMISION_DIRECTIVA: ['Presidente','Vicepresidente','Secretario','Prosecretario','Tesorero','Protesorero','Vocal Titular 1°','Vocal Titular 2°','Vocal Titular 3°','Vocal Suplente 1°','Vocal Suplente 2°','Vocal Suplente 3°'],
  REVISORA_CUENTAS: ['Revisor Titular','Revisor Suplente'],
  ARQUITECTURA: ['Miembro Arquitectura 1','Miembro Arquitectura 2','Miembro Arquitectura 3','Miembro Arquitectura 4','Miembro Arquitectura 5']
};
const NOMBRES_ORGANO = { COMISION_DIRECTIVA: 'Comisión Directiva', REVISORA_CUENTAS: 'Comisión Revisora de Cuentas', ARQUITECTURA: 'Comisión de Arquitectura' };

let PIN_SESION = '';

document.addEventListener('DOMContentLoaded', function() {
  actualizarCargos();
  const savedPin = localStorage.getItem('cgv_pin');
  if (!savedPin) return;
  PIN_SESION = savedPin;
  jsonp(EXEC_URL_PINES_CGV_CHECK + '?action=validarPin&pin=' + encodeURIComponent(savedPin) + '&modulo=gestion-institucional', function(res) {
    if (res && res.ok && res.valido) {
      document.getElementById('pinScreen').style.display = 'none';
      document.getElementById('mainScreen').style.display = 'block';
      document.getElementById('mainScreen').classList.add('active');
      cargarAutoridades();
    }
  });
});

function jsonp(url, cb) {
  const cbName = 'cb_' + Date.now() + Math.floor(Math.random()*1000);
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
      document.getElementById('pinScreen').style.display = 'none';
      document.getElementById('mainScreen').style.display = 'block';
      document.getElementById('mainScreen').classList.add('active');
      cargarAutoridades();
    } else {
      const input = document.getElementById('pinInput');
      input.classList.add('error');
      document.getElementById('pinError').textContent = 'PIN incorrecto';
      setTimeout(() => input.classList.remove('error'), 400);
    }
  });
}

function cerrarSesion() {
  localStorage.removeItem('cgv_pin');
  location.reload();
}

function actualizarCargos() {
  const organo = document.getElementById('fOrgano').value;
  const sel = document.getElementById('fCargo');
  sel.innerHTML = CARGOS_POR_ORGANO[organo].map(c => `<option value="${c}">${c}</option>`).join('');
}

function toggleForm() {
  document.getElementById('formPanel').classList.toggle('open');
}

function cargarAutoridades() {
  document.getElementById('loadingText').textContent = 'Cargando autoridades...';
  jsonp(EXEC_URL_GI + '?action=listarAutoridades', function(res) {
    document.getElementById('loadingText').textContent = '';
    if (!res || !res.ok) return;
    renderAutoridades(res.autoridades, res.vencenEsteAnio || []);
  });
}

function renderAutoridades(autoridades, vencenEsteAnio) {
  const cont = document.getElementById('listaOrganos');
  cont.innerHTML = '';

  const alerta = document.getElementById('resumenAlerta');
  if (vencenEsteAnio.length > 0) {
    alerta.style.display = 'block';
    alerta.textContent = '⚠️ Renuevan este año: ' + vencenEsteAnio.join(', ');
  } else {
    alerta.style.display = 'none';
  }

  ['COMISION_DIRECTIVA','REVISORA_CUENTAS','ARQUITECTURA'].forEach(organo => {
    const header = document.createElement('div');
    header.className = 'organo-header';
    header.textContent = NOMBRES_ORGANO[organo];
    cont.appendChild(header);

    const cargosDelOrgano = CARGOS_POR_ORGANO[organo];
    cargosDelOrgano.forEach(cargo => {
      const aut = autoridades.find(a => a.organo === organo && a.cargo === cargo && a.estado === 'VIGENTE');
      if (aut) {
        const div = document.createElement('div');
        div.className = 'cargo-card';
        const clase = aut.vence === 'este_anio' ? 'vence-este-anio' : (aut.vence === 'proximo_anio' ? 'vence-pronto' : 'vence-ok');
        const texto = aut.vence === 'este_anio' ? 'Renueva este año' : ('Vence ' + aut.fechaFinMandato);
        div.innerHTML = `<div class="cargo-info"><div class="cargo-nombre">${aut.nombre}</div><div class="cargo-tipo">${cargo}</div></div><span class="vence-badge ${clase}">${texto}</span>`;
        cont.appendChild(div);
      } else {
        const div = document.createElement('div');
        div.className = 'vacante-card';
        div.textContent = `${cargo} — VACANTE`;
        cont.appendChild(div);
      }
    });
  });
}

function guardarAutoridad() {
  const data = {
    organo: document.getElementById('fOrgano').value,
    cargo: document.getElementById('fCargo').value,
    nombre: document.getElementById('fNombre').value,
    dni: document.getElementById('fDni').value,
    fechaInicio: document.getElementById('fFechaInicio').value,
    pin: PIN_SESION
  };
  if (!data.nombre || !data.fechaInicio) { alert('Completá nombre y fecha de inicio'); return; }
  document.getElementById('loadingText').textContent = 'Guardando...';
  const params = Object.keys(data).map(k => k + '=' + encodeURIComponent(data[k])).join('&');
  jsonp(EXEC_URL_GI + '?action=guardarAutoridad&' + params, function(res) {
    document.getElementById('loadingText').textContent = '';
    if (res && res.ok) {
      toggleForm();
      document.getElementById('fNombre').value = '';
      document.getElementById('fDni').value = '';
      document.getElementById('fFechaInicio').value = '';
      cargarAutoridades();
    } else {
      alert('Error al guardar: ' + (res && res.error ? res.error : 'desconocido'));
    }
  });
}