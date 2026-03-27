// ═══════════════════════════════════════════════════════════════════════
// ESTADO
// ═══════════════════════════════════════════════════════════════════════
const estado = {
  titulo:      '',
  fechaInicio: null, // Date del momento en que se inicia el escrutinio
  candidatos:  [],   // [{ nombre: string, votos: number }]
  electores:   [],   // [{ nombre: string, votado: boolean }]
  historial:   [],   // [{ candidato: string, fechaHora: string }]
  fileHandle:  null, // File System Access API handle para guardar Excel
};

// ═══════════════════════════════════════════════════════════════════════
// INICIALIZACIÓN
// ═══════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('inp-excel').addEventListener('change', (e) => {
    if (e.target.files.length) cargarExcel(e.target.files[0]);
  });
  document.getElementById('btn-plantilla').addEventListener('click', descargarPlantilla);
  document.getElementById('btn-iniciar').addEventListener('click', iniciarEscrutinio);
  document.getElementById('btn-deshacer').addEventListener('click', deshacerVoto);
  document.getElementById('btn-exportar').addEventListener('click', exportarExcel);
  document.getElementById('btn-finalizar').addEventListener('click', finalizar);
  document.getElementById('btn-nueva').addEventListener('click', nuevaVotacion);
  document.getElementById('btn-excel-fin').addEventListener('click', exportarExcel);

  document.addEventListener('keydown', atajosTeclado);
});

// ═══════════════════════════════════════════════════════════════════════
// CARGA DEL EXCEL (candidatos + electores)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Dada una cabecera (array de strings), devuelve el índice de la columna
 * que más probablemente contiene el nombre de la persona.
 */
function detectarColNombre(cabecera) {
  const patron = /^(nombre|empleado|candidato|apellido|docente|profesor|name|teacher)/i;
  const idx = cabecera.findIndex(c => patron.test(String(c ?? '')));
  return idx !== -1 ? idx : 0;
}

/**
 * Dada una cabecera, devuelve el índice de la columna "Fecha de cese"
 * (para filtrar personal que ya no está en activo), o -1 si no existe.
 */
function detectarColCese(cabecera) {
  return cabecera.findIndex(c => /cese/i.test(String(c ?? '')));
}

/**
 * Devuelve true si la persona sigue en ejercicio, es decir:
 *   - La fecha de cese está vacía / ausente, O
 *   - La fecha de cese es igual o posterior a hoy (cese futuro o en curso).
 * Acepta tanto objetos Date (cuando SheetJS lee celdas de fecha nativas)
 * como cadenas de texto en formato DD/MM/YYYY, DD-MM-YYYY o ISO 8601.
 */
function esVigente(valorCese) {
  if (valorCese === null || valorCese === undefined) return true;
  const str = String(valorCese).trim();
  if (!str) return true;

  let fecha;
  if (valorCese instanceof Date) {
    fecha = new Date(valorCese); // copia para no mutar el original
  } else {
    // Formato DD/MM/YYYY o DD-MM-YYYY (habitual en Excel español)
    const mES = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (mES) {
      fecha = new Date(parseInt(mES[3]), parseInt(mES[2]) - 1, parseInt(mES[1]));
    } else {
      fecha = new Date(str); // ISO 8601 u otros formatos reconocidos por el motor JS
    }
  }

  if (isNaN(fecha?.getTime())) return true; // fecha no interpretable → incluir por seguridad

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  fecha.setHours(0, 0, 0, 0);

  return fecha >= hoy; // vigente si el cese es hoy o en el futuro
}

/**
 * Extrae nombres de una hoja SheetJS aplicando:
 *   - Detección automática de columna de nombre
 *   - Inclusión solo de filas cuya fecha de cese está vacía o no ha pasado
 */
function extraerNombres(ws) {
  // header:1 → array de arrays; primera fila = cabecera
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (filas.length < 2) return [];

  const cabecera  = filas[0].map(String);
  const colNombre = detectarColNombre(cabecera);
  const colCese   = detectarColCese(cabecera);

  return filas.slice(1)
    .filter(fila => {
      const nombre = String(fila[colNombre] ?? '').trim();
      if (!nombre) return false;
      if (colCese !== -1 && !esVigente(fila[colCese])) return false;
      return true;
    })
    .map(fila => String(fila[colNombre]).trim());
}

function cargarExcel(file) {
  const reader = new FileReader();

  reader.onload = (e) => {
    let wb;
    try {
      // cellDates: true → SheetJS convierte celdas de fecha nativas a objetos Date
      wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
    } catch {
      mostrarToast('No se pudo leer el fichero. ¿Es un Excel válido?', true);
      return;
    }

    // ── Localizar hoja de Candidatos ──────────────────────────────────
    const nombreHojaCand = wb.SheetNames.find(n => /candidat/i.test(n));
    if (!nombreHojaCand) {
      mostrarToast('No se encontró la hoja "Candidatos" en el Excel', true);
      return;
    }

    // ── Localizar hoja de Electores ───────────────────────────────────
    const nombreHojaElec = wb.SheetNames.find(
      n => /elector|votant|docent|profesor|emplead/i.test(n)
    );
    if (!nombreHojaElec) {
      mostrarToast('No se encontró la hoja "Electores" en el Excel', true);
      return;
    }

    if (nombreHojaCand === nombreHojaElec) {
      mostrarToast('Las hojas "Candidatos" y "Electores" deben ser distintas', true);
      return;
    }

    // ── Extraer datos ─────────────────────────────────────────────────
    const nombresCand = extraerNombres(wb.Sheets[nombreHojaCand]);
    const nombresElec = extraerNombres(wb.Sheets[nombreHojaElec]);

    if (nombresCand.length === 0) {
      mostrarToast('La hoja "Candidatos" no contiene nombres válidos', true);
      return;
    }
    if (nombresElec.length === 0) {
      mostrarToast('La hoja "Electores" no contiene nombres válidos', true);
      return;
    }

    estado.candidatos = nombresCand.map(nombre => ({ nombre, votos: 0 }));
    estado.electores  = nombresElec.map(nombre => ({ nombre, votado: false }));

    mostrarPrevia();
    mostrarToast(`${nombresCand.length} candidatos · ${nombresElec.length} electores cargados`);
  };

  reader.onerror = () => mostrarToast('Error al leer el fichero', true);
  reader.readAsArrayBuffer(file);
}

// ── Vista previa en pantalla de configuración ─────────────────────────
function mostrarPrevia() {
  const MAX_VISIBLE = 6;

  function renderBloque(contenedorId, titulo, items) {
    const bloque = document.getElementById(contenedorId);
    const extra  = items.length - MAX_VISIBLE;
    bloque.innerHTML = `
      <h4>${titulo} <span class="badge">${items.length}</span></h4>
      <ul class="preview-lista">
        ${items.slice(0, MAX_VISIBLE).map(n => `<li>${n}</li>`).join('')}
      </ul>
      ${extra > 0 ? `<p class="preview-mas">… y ${extra} más</p>` : ''}
    `;
  }

  renderBloque('prev-candidatos', 'Candidatos', estado.candidatos.map(c => c.nombre));
  renderBloque('prev-electores',  'Electores',  estado.electores.map(e => e.nombre));

  document.getElementById('preview').style.display = '';
}

// ═══════════════════════════════════════════════════════════════════════
// PLANTILLA EXCEL DESCARGABLE
// ═══════════════════════════════════════════════════════════════════════
function descargarPlantilla() {
  const wb = XLSX.utils.book_new();

  // Candidatos: misma estructura que Electores; se filtra por Fecha de cese igual
  const wsCand = XLSX.utils.aoa_to_sheet([
    ['Nombre', 'Fecha de cese'],
    ['García López, Ana',          ''],           // sin cese → incluida
    ['Martínez Ruiz, Carlos',      ''],           // sin cese → incluido
    ['Fernández Jiménez, María',   ''],           // sin cese → incluida
    ['Sánchez Romero, Luis',       '31/01/2025'], // cese pasado → excluido
  ]);
  wsCand['!cols'] = [{ wch: 36 }, { wch: 16 }];

  const wsElec = XLSX.utils.aoa_to_sheet([
    ['Nombre', 'Fecha de cese'],
    ['Álvarez Moreno, Carmen',    ''],           // sin cese → incluida
    ['Blanco Ruiz, Javier',       ''],           // sin cese → incluido
    ['Castro López, Dolores',     ''],           // sin cese → incluida
    ['Díaz Fernández, Antonio',   '30/06/2025'], // cese futuro → incluido
    ['Escribano García, Lucía',   '15/01/2025'], // cese pasado → excluida
  ]);
  wsElec['!cols'] = [{ wch: 36 }, { wch: 16 }];

  XLSX.utils.book_append_sheet(wb, wsCand, 'Candidatos');
  XLSX.utils.book_append_sheet(wb, wsElec, 'Electores');

  XLSX.writeFile(wb, 'plantilla_votacion.xlsx');
  mostrarToast('Plantilla descargada');
}

// ═══════════════════════════════════════════════════════════════════════
// INICIO DEL ESCRUTINIO
// ═══════════════════════════════════════════════════════════════════════
function iniciarEscrutinio() {
  const titulo = document.getElementById('inp-titulo').value.trim();
  if (!titulo) { mostrarToast('Escribe el título de la votación', true); return; }
  if (estado.candidatos.length < 2) {
    mostrarToast('Carga el fichero Excel con al menos 2 candidatos', true);
    return;
  }

  // Reiniciar contadores (por si se vuelve atrás)
  estado.candidatos.forEach(c => { c.votos = 0; });
  estado.electores.forEach(e  => { e.votado = false; });
  estado.historial  = [];
  estado.fileHandle = null;
  estado.titulo     = titulo;
  estado.fechaInicio = new Date();

  document.getElementById('esc-titulo').textContent = titulo;

  if (estado.electores.length > 0) {
    document.getElementById('stat-pendientes-wrap').style.display = '';
    document.getElementById('stat-electores-wrap').style.display  = '';
    document.getElementById('electores-panel').style.display = '';
    renderizarListaElectores();
  }

  renderizarBotonesVoto();
  renderizarResultados();
  actualizarEstadisticas();
  irA('esc');
}

// ═══════════════════════════════════════════════════════════════════════
// REGISTRO DE VOTOS
// ═══════════════════════════════════════════════════════════════════════
function renderizarBotonesVoto() {
  const cont = document.getElementById('botones-voto');
  cont.innerHTML = '';
  estado.candidatos.forEach((c, i) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn-voto';
    btn.textContent = c.nombre;
    btn.addEventListener('click', () => registrarVoto(i));
    cont.appendChild(btn);
  });
}

function registrarVoto(idx) {
  const c    = estado.candidatos[idx];
  const ahora = new Date();
  c.votos++;
  estado.historial.push({
    candidato: c.nombre,
    fechaHora: ahora.toLocaleDateString('es-ES') + ' ' + ahora.toLocaleTimeString('es-ES'),
  });
  actualizarEstadisticas();
  renderizarResultados(idx);
}

function deshacerVoto() {
  if (!estado.historial.length) { mostrarToast('No hay votos que deshacer', true); return; }
  const ultimo = estado.historial.pop();
  const c = estado.candidatos.find(x => x.nombre === ultimo.candidato);
  if (c && c.votos > 0) c.votos--;
  actualizarEstadisticas();
  renderizarResultados();
  mostrarToast(`Voto de "${ultimo.candidato}" deshecho`);
}

function totalVotos() {
  return estado.candidatos.reduce((s, c) => s + c.votos, 0);
}

function actualizarEstadisticas() {
  const emitidos = totalVotos();
  document.getElementById('stat-total').textContent = emitidos;

  if (estado.electores.length > 0) {
    const votados = estado.electores.filter(e => e.votado).length;
    const pend    = Math.max(0, estado.electores.length - emitidos);
    document.getElementById('stat-pendientes').textContent = pend;
    document.getElementById('stat-electores-val').textContent = `${votados}/${estado.electores.length}`;
    document.getElementById('electores-count').textContent    = `${votados}/${estado.electores.length}`;
  }
}

// ═══════════════════════════════════════════════════════════════════════
// LISTA DE ELECTORES
// ═══════════════════════════════════════════════════════════════════════
function renderizarListaElectores() {
  const cont = document.getElementById('lista-electores');
  cont.innerHTML = '';
  estado.electores.forEach((elector, i) => {
    const lbl = document.createElement('label');
    lbl.className = 'elector-row' + (elector.votado ? ' votado' : '');

    const chk = document.createElement('input');
    chk.type    = 'checkbox';
    chk.checked = elector.votado;
    chk.addEventListener('change', () => {
      estado.electores[i].votado = chk.checked;
      lbl.classList.toggle('votado', chk.checked);
      actualizarEstadisticas();
    });

    const span = document.createElement('span');
    span.textContent = elector.nombre;

    lbl.appendChild(chk);
    lbl.appendChild(span);
    cont.appendChild(lbl);
  });
}

// ═══════════════════════════════════════════════════════════════════════
// RESULTADOS EN VIVO
// ═══════════════════════════════════════════════════════════════════════
function renderizarResultados(idxFlash = -1) {
  const total    = totalVotos();
  const maxVotos = estado.candidatos.reduce((m, c) => Math.max(m, c.votos), 1);
  const panel    = document.getElementById('panel-resultados');
  const lideres  = estado.candidatos.reduce((m, c) => Math.max(m, c.votos), 0);

  panel.innerHTML = '';
  estado.candidatos.forEach((c, i) => {
    const pct   = total > 0 ? (c.votos / total * 100).toFixed(1) : '0.0';
    const barra = (c.votos / maxVotos * 100).toFixed(1);
    const esLider = c.votos > 0 && c.votos === lideres;

    const card = document.createElement('div');
    card.className = 'candidato-card' + (esLider ? ' ganador' : '');
    card.innerHTML = `
      <div class="cand-nombre">${c.nombre}</div>
      <div class="cand-votos">${c.votos}</div>
      <div class="cand-barra-wrap">
        <div class="cand-barra" style="width:${barra}%"></div>
      </div>
      <div class="cand-pct">${pct}%</div>
    `;
    panel.appendChild(card);

    if (i === idxFlash) {
      requestAnimationFrame(() => {
        card.classList.add('flash');
        card.addEventListener('animationend', () => card.classList.remove('flash'), { once: true });
      });
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// FINALIZAR
// ═══════════════════════════════════════════════════════════════════════
function finalizar() {
  if (!totalVotos()) { mostrarToast('No se ha registrado ningún voto aún', true); return; }
  if (!confirm('¿Finalizar el escrutinio y ver los resultados definitivos?')) return;

  const total  = totalVotos();
  const sorted = [...estado.candidatos].sort((a, b) => b.votos - a.votos);

  document.getElementById('fin-subtitulo').textContent =
    `${estado.titulo} — ${total} votos escrutados`;

  const tabla = document.getElementById('fin-tabla');
  tabla.innerHTML = '';
  sorted.forEach((c, i) => {
    const pct  = (c.votos / total * 100).toFixed(1);
    const fila = document.createElement('div');
    fila.className = 'fila-final' + (i === 0 ? ' primero' : '');
    fila.innerHTML = `
      <div class="pos">${i === 0 ? '🏆' : i + 1}</div>
      <div class="nombre">${c.nombre}</div>
      <div class="votos">${c.votos}</div>
      <div class="pct">${pct}%</div>
    `;
    tabla.appendChild(fila);
  });

  irA('fin');
}

function nuevaVotacion() {
  if (!confirm('¿Iniciar una nueva votación? Se perderán los datos actuales si no los has exportado.')) return;

  estado.candidatos = [];
  estado.electores  = [];
  estado.historial  = [];
  estado.titulo     = '';
  estado.fileHandle = null;

  document.getElementById('inp-titulo').value = '';
  document.getElementById('inp-excel').value  = '';
  document.getElementById('preview').style.display = 'none';
  document.getElementById('stat-pendientes-wrap').style.display = 'none';
  document.getElementById('stat-electores-wrap').style.display  = 'none';
  document.getElementById('electores-panel').style.display = 'none';

  irA('cfg');
}

// ═══════════════════════════════════════════════════════════════════════
// EXPORTAR RESULTADOS A EXCEL
// ═══════════════════════════════════════════════════════════════════════
function construirLibroResultados() {
  const total  = totalVotos();
  const sorted = [...estado.candidatos].sort((a, b) => b.votos - a.votos);

  const fmt = (d) => d.toLocaleDateString('es-ES') + ' ' + d.toLocaleTimeString('es-ES');

  // Hoja 1 — Resumen
  const filas = [
    ['Votación:', estado.titulo],
    ['Inicio del escrutinio:', estado.fechaInicio ? fmt(estado.fechaInicio) : '—'],
    ['Exportación de resultados:', fmt(new Date())],
    ['Total votos escrutados:', total],
    ['Total electores:', estado.electores.length || '—'],
    [],
    ['Posición', 'Candidato', 'Votos', '% sobre total'],
    ...sorted.map((c, i) => [
      i + 1,
      c.nombre,
      c.votos,
      total > 0 ? parseFloat((c.votos / total * 100).toFixed(2)) : 0,
    ]),
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(filas);
  ws1['!cols'] = [{ wch: 22 }, { wch: 36 }, { wch: 8 }, { wch: 14 }];

  // Hoja 2 — Detalle papeleta a papeleta
  const detalle = [
    ['N°', 'Candidato', 'Fecha y hora de registro'],
    ...estado.historial.map((h, i) => [i + 1, h.candidato, h.fechaHora]),
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(detalle);
  ws2['!cols'] = [{ wch: 5 }, { wch: 36 }, { wch: 16 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, 'Resumen');
  XLSX.utils.book_append_sheet(wb, ws2, 'Detalle papeletas');

  // Hoja 3 — Estado de participación de electores (si se cargaron)
  if (estado.electores.length > 0) {
    const wsElec = XLSX.utils.aoa_to_sheet([
      ['Elector', 'Ha votado'],
      ...estado.electores.map(e => [e.nombre, e.votado ? 'Sí' : 'No']),
    ]);
    wsElec['!cols'] = [{ wch: 36 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, wsElec, 'Participación');
  }

  return wb;
}

async function exportarExcel() {
  if (!totalVotos()) { mostrarToast('No hay votos registrados aún', true); return; }

  const wb    = construirLibroResultados();
  const slug  = estado.titulo.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '').slice(0, 40);
  const fname = `escrutinio_${slug}.xlsx`;

  // Preferir File System Access API (Chrome/Edge)
  if ('showSaveFilePicker' in window) {
    try {
      if (!estado.fileHandle) {
        estado.fileHandle = await window.showSaveFilePicker({
          suggestedName: fname,
          types: [{
            description: 'Libro Excel',
            accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
          }],
        });
      }
      const buf      = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const writable = await estado.fileHandle.createWritable();
      await writable.write(new Blob([buf]));
      await writable.close();
      mostrarToast('Resultados guardados en Excel');
      return;
    } catch (err) {
      if (err.name === 'AbortError') return;
      estado.fileHandle = null;
    }
  }

  // Fallback: descarga directa (Firefox, Safari…)
  XLSX.writeFile(wb, fname);
  mostrarToast('Resultados descargados como Excel');
}

// ═══════════════════════════════════════════════════════════════════════
// UTILIDADES
// ═══════════════════════════════════════════════════════════════════════
function irA(id) {
  document.querySelectorAll('.pantalla').forEach(p => p.classList.remove('activa'));
  document.getElementById(id).classList.add('activa');
}

let toastTimer;
function mostrarToast(msg, error = false) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className   = 'visible' + (error ? ' error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.className = ''; }, 2800);
}

// Atajos de teclado durante el escrutinio:
//   1–9  → voto al candidato N
//   Ctrl+Z → deshacer
function atajosTeclado(e) {
  if (!document.getElementById('esc').classList.contains('activa')) return;
  if (e.target.tagName === 'INPUT') return;

  const n = parseInt(e.key);
  if (n >= 1 && n <= estado.candidatos.length) { registrarVoto(n - 1); return; }
  if (e.key === 'z' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); deshacerVoto(); }
}
