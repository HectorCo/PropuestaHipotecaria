(() => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel } = window.docx;

  // ---------- Referencias DOM ----------
  const form = document.getElementById('formPropuesta');
  const modal = document.getElementById('previewModal');
  const previewContent = document.getElementById('previewContent');
  const btnPreview = document.getElementById('btnPreview');
  const btnLimpiar = document.getElementById('btnLimpiar');
  const btnGenerar = document.getElementById('btnGenerar');
  const btnClose = document.getElementById('btnClose');

  const bloqueMixta = document.getElementById('bloqueMixta');
  const selectTipoInteres = form.elements.tipoInteres;
  const listaBonificaciones = document.getElementById('listaBonificaciones');
  const tbodyGastos = document.getElementById('tbodyGastos');

  // Borradores
  const draftStatus = document.getElementById('draftStatus');
  const btnNuevo = document.getElementById('btnNuevo');
  const btnGuardar = document.getElementById('btnGuardar');
  const btnGuardarComo = document.getElementById('btnGuardarComo');
  const btnMisBorradores = document.getElementById('btnMisBorradores');

  const draftsModal = document.getElementById('draftsModal');
  const listaDrafts = document.getElementById('listaDrafts');
  const btnImportar = document.getElementById('btnImportar');
  const btnExportarTodos = document.getElementById('btnExportarTodos');
  const inputImportar = document.getElementById('inputImportar');

  const STORAGE_KEY = 'propuesta_scratch_v2';
  const DRAFTS_KEY = 'propuesta_drafts_v2';
  const FONT_DOCX = 'Aptos, Calibri, Segoe UI, sans-serif';

  const FILAS_GASTOS_DEFAULT = [
    'Registro de la propiedad',
    'Notaría',
    'Gestoría',
    'Impuestos',
    'Tasación',
    'Nota simple',
    'Seguro de daños',
  ];

  let ultimosDatos = null;
  let currentDraftId = null;
  let currentDraftName = '';

  // ---------- Utilidades ----------
  const eur = (v) => {
    const s = String(v ?? '').trim();
    if (!s) return '—';
    if (s.includes('€')) return s;
    const n = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    if (isNaN(n)) return s + ' €';
    return new Intl.NumberFormat('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n) + ' €';
  };

  const pct = (v) => {
    const s = String(v ?? '').trim();
    if (!s) return '—';
    if (s.includes('%')) return s;
    return s + '%';
  };

  const fmtFecha = (iso) => {
    if (!iso) return '';
    const d = new Date(iso + 'T00:00:00');
    return d.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
  };

  const fmtFechaHora = (iso) => {
    if (!iso) return '';
    const d = new Date(iso);
    return d.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
           ' ' + d.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' });
  };

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const generarId = () => {
    if (crypto?.randomUUID) return crypto.randomUUID();
    return 'id-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
  };

  const slugify = (s) => String(s || 'borrador')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 60) || 'borrador';

  // ---------- Filas dinámicas: bonificaciones ----------
  function crearFilaBonificacion(descripcion = '', puntos = '') {
    const fila = document.createElement('div');
    fila.className = 'fila-dinamica';
    fila.innerHTML = `
      <input type="text" name="bonif_desc[]" placeholder="Ej. Domiciliación de nómina" value="${esc(descripcion)}" />
      <input type="text" name="bonif_puntos[]" placeholder="0,50 p.p." value="${esc(puntos)}" />
      <button type="button" class="btn-remove" aria-label="Eliminar">×</button>
    `;
    fila.querySelector('.btn-remove').addEventListener('click', () => {
      fila.remove();
      autosave();
    });
    return fila;
  }

  function agregarBonificacion(datos) {
    listaBonificaciones.appendChild(crearFilaBonificacion(datos?.descripcion, datos?.puntos));
  }

  // ---------- Tabla dinámica: gastos ----------
  function crearFilaGasto(concepto = '', cliente = '', entidad = '') {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" name="gasto_concepto[]" value="${esc(concepto)}" placeholder="Concepto" /></td>
      <td><input type="text" name="gasto_cliente[]" value="${esc(cliente)}" placeholder="0,00 €" /></td>
      <td><input type="text" name="gasto_entidad[]" value="${esc(entidad)}" placeholder="0,00 €" /></td>
      <td><button type="button" class="btn-remove" aria-label="Eliminar">×</button></td>
    `;
    tr.querySelector('.btn-remove').addEventListener('click', () => {
      tr.remove();
      autosave();
    });
    return tr;
  }

  function agregarGasto(concepto, cliente, entidad) {
    tbodyGastos.appendChild(crearFilaGasto(concepto, cliente, entidad));
  }

  function inicializarGastos() {
    tbodyGastos.innerHTML = '';
    FILAS_GASTOS_DEFAULT.forEach((c) => agregarGasto(c, '', ''));
  }

  function inicializarBonificaciones() {
    listaBonificaciones.innerHTML = '';
    agregarBonificacion({ descripcion: 'Domiciliación de nómina', puntos: '0,50 p.p.' });
    agregarBonificacion({ descripcion: 'Seguro de hogar', puntos: '0,20 p.p.' });
  }

  // ---------- Lectura / escritura del formulario ----------
  function leerDatos() {
    const d = {};
    for (const el of form.elements) {
      if (el.name && !el.name.endsWith('[]')) d[el.name] = (el.value || '').trim();
    }
    d._bonificaciones = [...listaBonificaciones.querySelectorAll('.fila-dinamica')].map((fila) => ({
      descripcion: fila.querySelector('input[name="bonif_desc[]"]')?.value.trim() || '',
      puntos: fila.querySelector('input[name="bonif_puntos[]"]')?.value.trim() || '',
    })).filter((b) => b.descripcion);
    d._gastos = [...tbodyGastos.querySelectorAll('tr')].map((tr) => ({
      concepto: tr.querySelector('input[name="gasto_concepto[]"]')?.value.trim() || '',
      cliente: tr.querySelector('input[name="gasto_cliente[]"]')?.value.trim() || '',
      entidad: tr.querySelector('input[name="gasto_entidad[]"]')?.value.trim() || '',
    })).filter((g) => g.concepto || g.cliente || g.entidad);
    return d;
  }

  function aplicarDatosAlFormulario(datos) {
    if (!datos) datos = {};
    form.reset();

    for (const [k, v] of Object.entries(datos)) {
      if (k.startsWith('_')) continue;
      const el = form.elements[k];
      if (el) el.value = v ?? '';
    }

    listaBonificaciones.innerHTML = '';
    const bonif = Array.isArray(datos._bonificaciones) ? datos._bonificaciones : [];
    if (bonif.length) bonif.forEach(agregarBonificacion);
    else inicializarBonificaciones();

    tbodyGastos.innerHTML = '';
    const gastos = Array.isArray(datos._gastos) ? datos._gastos : [];
    if (gastos.length) gastos.forEach((g) => agregarGasto(g.concepto, g.cliente, g.entidad));
    else inicializarGastos();

    actualizarBloqueMixta();
  }

  // ---------- Bloque "Mixta" ----------
  function actualizarBloqueMixta() {
    const esMixta = selectTipoInteres.value === 'Mixta';
    bloqueMixta.hidden = !esMixta;
    if (!esMixta) {
      ['periodoFijoAnos', 'tinFijo', 'diferencialVariable'].forEach((n) => {
        if (form.elements[n]) form.elements[n].value = '';
      });
      if (form.elements.indiceVariable) form.elements.indiceVariable.value = 'Euríbor 12M';
    }
  }

  // =========================================================
  // ================  GESTIÓN DE BORRADORES  ================
  // =========================================================
  function getDrafts() {
    try { return JSON.parse(localStorage.getItem(DRAFTS_KEY) || '[]'); }
    catch { return []; }
  }
  function saveDrafts(arr) {
    try { localStorage.setItem(DRAFTS_KEY, JSON.stringify(arr)); } catch (_) {}
  }
  function findDraft(id) {
    return getDrafts().find((d) => d.id === id);
  }
  function upsertDraft(draft) {
    const drafts = getDrafts();
    const i = drafts.findIndex((d) => d.id === draft.id);
    if (i >= 0) drafts[i] = draft;
    else drafts.push(draft);
    saveDrafts(drafts);
  }
  function removeDraft(id) {
    saveDrafts(getDrafts().filter((d) => d.id !== id));
  }

  function actualizarBarra() {
    if (currentDraftId && currentDraftName) {
      draftStatus.textContent = currentDraftName;
      draftStatus.classList.remove('unsaved');
    } else {
      draftStatus.textContent = 'Sin guardar';
      draftStatus.classList.add('unsaved');
    }
  }

  // ---------- Autosave ----------
  function autosave() {
    const datos = leerDatos();
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(datos)); } catch (_) {}

    if (currentDraftId) {
      const draft = findDraft(currentDraftId);
      if (draft) {
        draft.datos = datos;
        draft.updatedAt = new Date().toISOString();
        upsertDraft(draft);
      }
    }
  }

  function cargarScratch() {
    try {
      const datos = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
      aplicarDatosAlFormulario(datos);
    } catch (_) {
      aplicarDatosAlFormulario({});
    }
  }

  // ---------- Guardar ----------
  function guardarBorradorActual() {
    if (!currentDraftId) {
      guardarComoNuevoBorrador();
      return;
    }
    const datos = leerDatos();
    const draft = findDraft(currentDraftId);
    if (!draft) {
      guardarComoNuevoBorrador();
      return;
    }
    draft.datos = datos;
    draft.updatedAt = new Date().toISOString();
    if (!draft.nombre && datos.nombreCliente) draft.nombre = datos.nombreCliente;
    upsertDraft(draft);
    currentDraftName = draft.nombre;
    actualizarBarra();
  }

  function guardarComoNuevoBorrador() {
    const datos = leerDatos();
    const sugerido = datos.nombreCliente || 'Nuevo borrador';
    const nombre = prompt('Nombre del borrador:', sugerido);
    if (nombre === null) return;
    const nombreFinal = (nombre || '').trim() || 'Sin nombre';
    const id = generarId();
    const now = new Date().toISOString();
    const draft = {
      id,
      nombre: nombreFinal,
      createdAt: now,
      updatedAt: now,
      datos,
    };
    upsertDraft(draft);
    currentDraftId = id;
    currentDraftName = nombreFinal;
    actualizarBarra();
  }

  // ---------- Nuevo ----------
  function nuevoBorrador() {
    if (!confirm('¿Crear un borrador nuevo? Los datos no guardados se perderán.')) return;
    currentDraftId = null;
    currentDraftName = '';
    localStorage.removeItem(STORAGE_KEY);
    form.reset();
    form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    inicializarBonificaciones();
    inicializarGastos();
    actualizarBloqueMixta();
    actualizarBarra();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ---------- Abrir ----------
  function abrirBorrador(id) {
    const draft = findDraft(id);
    if (!draft) { alert('Borrador no encontrado.'); return; }
    aplicarDatosAlFormulario(draft.datos || {});
    currentDraftId = draft.id;
    currentDraftName = draft.nombre;
    actualizarBarra();
    cerrarDraftsModal();
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(draft.datos || {})); } catch (_) {}
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ---------- Eliminar ----------
  function eliminarBorrador(id) {
    const draft = findDraft(id);
    if (!draft) return;
    if (!confirm(`¿Eliminar el borrador "${draft.nombre}"? Esta acción no se puede deshacer.`)) return;
    removeDraft(id);
    if (currentDraftId === id) {
      currentDraftId = null;
      currentDraftName = '';
      actualizarBarra();
    }
    renderizarListaDrafts();
  }

  // ---------- Renombrar ----------
  function renombrarBorrador(id) {
    const draft = findDraft(id);
    if (!draft) return;
    const nuevo = prompt('Nuevo nombre del borrador:', draft.nombre);
    if (nuevo === null) return;
    const nombreFinal = (nuevo || '').trim();
    if (!nombreFinal) { alert('El nombre no puede estar vacío.'); return; }
    draft.nombre = nombreFinal;
    draft.updatedAt = new Date().toISOString();
    upsertDraft(draft);
    if (currentDraftId === id) {
      currentDraftName = nombreFinal;
      actualizarBarra();
    }
    renderizarListaDrafts();
  }

  // ---------- Duplicar ----------
  function duplicarBorrador(id) {
    const draft = findDraft(id);
    if (!draft) return;
    const copia = JSON.parse(JSON.stringify(draft));
    copia.id = generarId();
    copia.nombre = draft.nombre + ' (copia)';
    copia.createdAt = new Date().toISOString();
    copia.updatedAt = new Date().toISOString();
    upsertDraft(copia);
    renderizarListaDrafts();
  }

  // ---------- Exportar ----------
  function exportarBorrador(id) {
    const draft = findDraft(id);
    if (!draft) return;
    const blob = new Blob([JSON.stringify(draft, null, 2)], { type: 'application/json' });
    const nombre = `borrador_${slugify(draft.nombre)}_${new Date().toISOString().slice(0, 10)}.json`;
    descargarBlob(blob, nombre);
  }

  function exportarTodos() {
    const drafts = getDrafts();
    if (!drafts.length) { alert('No hay borradores para exportar.'); return; }
    const blob = new Blob([JSON.stringify(drafts, null, 2)], { type: 'application/json' });
    const nombre = `borradores_${new Date().toISOString().slice(0, 10)}.json`;
    descargarBlob(blob, nombre);
  }

  // ---------- Importar ----------
  function importarBorradores(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        const lista = Array.isArray(data) ? data : [data];
        const drafts = getDrafts();
        let añadidos = 0;
        lista.forEach((item) => {
          if (!item || typeof item !== 'object' || !item.datos) return;
          const copia = JSON.parse(JSON.stringify(item));
          copia.id = generarId();
          copia.nombre = (copia.nombre || 'Importado').trim() || 'Importado';
          copia.createdAt = copia.createdAt || new Date().toISOString();
          copia.updatedAt = new Date().toISOString();
          copia.importedAt = new Date().toISOString();
          drafts.push(copia);
          añadidos++;
        });
        saveDrafts(drafts);
        renderizarListaDrafts();
        if (añadidos === 0) {
          alert('No se encontraron borradores válidos en el archivo.');
        } else {
          alert(`${añadidos} borrador(es) importado(s) correctamente.`);
        }
      } catch (err) {
        console.error(err);
        alert('Error al importar el archivo: ' + err.message);
      }
    };
    reader.onerror = () => alert('Error al leer el archivo.');
    reader.readAsText(file);
  }

  // ---------- Render lista ----------
  function renderizarListaDrafts() {
    const drafts = getDrafts().sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
    if (!drafts.length) {
      listaDrafts.innerHTML = `
        <div class="draft-empty">
          No tienes borradores guardados todavía.<br>
          Pulsa "Guardar como…" para crear uno, o "Importar JSON" para cargar uno existente.
        </div>`;
      return;
    }
    listaDrafts.innerHTML = drafts.map((d) => `
      <div class="draft-item${d.id === currentDraftId ? ' active' : ''}" data-id="${esc(d.id)}">
        <div class="draft-item-info">
          <div class="draft-item-name">${esc(d.nombre)}</div>
          <div class="draft-item-meta">
            ${d.datos?.nombreCliente ? esc(d.datos.nombreCliente) + ' · ' : ''}
            Actualizado: ${esc(fmtFechaHora(d.updatedAt))}
          </div>
        </div>
        <div class="draft-item-actions">
          <button type="button" class="btn-ghost" data-action="abrir">Abrir</button>
          <button type="button" class="btn-ghost" data-action="renombrar">Renombrar</button>
          <button type="button" class="btn-ghost" data-action="duplicar">Duplicar</button>
          <button type="button" class="btn-ghost" data-action="exportar">Exportar</button>
          <button type="button" class="btn-danger" data-action="eliminar">Eliminar</button>
        </div>
      </div>
    `).join('');

    listaDrafts.querySelectorAll('.draft-item').forEach((el) => {
      const id = el.dataset.id;
      el.querySelectorAll('button[data-action]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const action = btn.dataset.action;
          if (action === 'abrir') abrirBorrador(id);
          else if (action === 'renombrar') renombrarBorrador(id);
          else if (action === 'duplicar') duplicarBorrador(id);
          else if (action === 'exportar') exportarBorrador(id);
          else if (action === 'eliminar') eliminarBorrador(id);
        });
      });
    });
  }

  // ---------- Modal borradores ----------
  function abrirDraftsModal() {
    renderizarListaDrafts();
    draftsModal.hidden = false;
    document.body.style.overflow = 'hidden';
  }
  function cerrarDraftsModal() {
    draftsModal.hidden = true;
    document.body.style.overflow = '';
  }

  // =========================================================
  // ================  TEXTO DEL DOCUMENTO  ==================
  // =========================================================
  function textos(d) {
    const conclusionDefault =
      `Esta propuesta equilibra una cuota mensual cómoda con una financiación ajustada a tus necesidades. ` +
      `Es una oportunidad sólida para adquirir tu hogar con total seguridad financiera.`;

    const notasLegalesDefault =
      `Documento informativo no vinculante. Las condiciones aquí reflejadas son una simulación y pueden ` +
      `variar según las características de los interesados, del inmueble a hipotecar y de las condiciones ` +
      `del mercado vigentes en cada momento. En caso de formalizarse la operación y no efectuar los pagos ` +
      `puntualmente, se responderá no solo con el inmueble hipotecado sino también con todos los bienes y ` +
      `derechos presentes y futuros. Este documento no sustituye ni modifica la información precontractual ` +
      `legalmente prevista.`;

    let condiciones;
    let detalleCondiciones = [];

    if (d.tipoInteres === 'Mixta') {
      const aniosFijos = d.periodoFijoAnos || '—';
      const aniosVariables = (() => {
        const tot = parseInt(d.plazoAnos, 10);
        const fij = parseInt(d.periodoFijoAnos, 10);
        return (!isNaN(tot) && !isNaN(fij) && tot > fij) ? String(tot - fij) : '—';
      })();
      condiciones =
        `Modalidad Mixta: un primer periodo a tipo fijo que te protege frente a las subidas de tipos, ` +
        `seguido de un periodo a tipo variable referenciado al ${d.indiceVariable || 'Euríbor 12M'} más un diferencial.`;
      detalleCondiciones = [
        `Periodo fijo (${aniosFijos} años): TIN fijo al ${pct(d.tinFijo)}.`,
        `Periodo variable (${aniosVariables} años): ${d.indiceVariable || 'Euríbor 12M'} + ${pct(d.diferencialVariable)}.`,
        `Plazo total: ${d.plazoAnos || '—'} años.`,
      ];
    } else {
      condiciones = `Modalidad ${d.tipoInteres || 'Fija'} durante toda la vida del préstamo.`;
      detalleCondiciones = [`Plazo: ${d.plazoAnos || '—'} años.`];
    }

    return {
      condiciones,
      detalleCondiciones,
      conclusion: (d.textoConclusion || '').trim() || conclusionDefault,
      notasLegales: (d.notasLegales || '').trim() || notasLegalesDefault,
      textoAhorro: (d.textoAhorro || '').trim() ||
        '(Precio de compra − Importe financiado + comisiones y gastos a su cargo)',
      subtitle: [
        d.nombreCliente,
        d.dniCliente,
        d.fechaDocumento ? fmtFecha(d.fechaDocumento) : '',
      ].filter(Boolean).join(' · '),
    };
  }

  // =========================================================
  // ================  VISTA PREVIA HTML  ====================
  // =========================================================
  function construirHTML(d) {
    const t = textos(d);
    const row = (k, v) => `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`;

    const filasDatosOperacion = [
      row('Precio de compra', eur(d.precioCompra)),
      row('Importe del préstamo hipotecario', eur(d.capitalHipotecario)),
      row('Plazo', (d.plazoAnos || '—') + ' años'),
      row('Tipo elegido', d.tipoInteres || '—'),
      row('Producto comercial', d.productoComercial || '—'),
      row('Finalidad', d.finalidad || '—'),
      row('Tipo de inmueble', d.tipoInmueble || '—'),
      row('Ubicación del inmueble', d.ubicacionInmueble || '—'),
    ].join('');

    const filasCondiciones = [
      row('Cuota mensual — con bonificación', eur(d.cuotaBonificada)),
      row('Cuota mensual — sin bonificación', eur(d.cuotaSinBonificar)),
      row('TIN — con bonificación', pct(d.tinBonificado)),
      row('TIN — sin bonificación', pct(d.tinSinBonificar)),
      row('TAE — con bonificación', pct(d.taeBonificada)),
      row('TAE — sin bonificación', pct(d.taeSinBonificar)),
      row('Importe total adeudado — con bonificación', eur(d.importeTotalBonificado)),
      row('Importe total adeudado — sin bonificación', eur(d.importeTotalSinBonificar)),
      d.numeroCuotas ? row('Número de cuotas', d.numeroCuotas) : '',
    ].filter(Boolean).join('');

    const filasBonificaciones = d._bonificaciones.map((b) =>
      `<tr><td>${esc(b.descripcion)}</td><td class="num">${esc(b.puntos || '—')}</td></tr>`
    ).join('');

    const filasGastos = d._gastos.map((g) =>
      `<tr>
        <td>${esc(g.concepto)}</td>
        <td class="num">${g.cliente ? esc(eur(g.cliente)) : '—'}</td>
        <td class="num">${g.entidad ? esc(eur(g.entidad)) : '—'}</td>
      </tr>`
    ).join('');

    return `
      <h1 class="doc-title">Propuesta de Financiación Hipotecaria</h1>
      <p class="doc-subtitle">${esc(t.subtitle)}</p>

      <table class="doc-table">
        ${d.oficina ? row('Oficina', d.oficina) : ''}
        ${d.gestor ? row('Gestor', d.gestor) : ''}
        ${d.expediente ? row('Expediente', d.expediente) : ''}
        ${d.emailGestor ? row('Email gestor', d.emailGestor) : ''}
        ${d.telefonoGestor ? row('Teléfono gestor', d.telefonoGestor) : ''}
      </table>

      <h2>1. Datos de la operación</h2>
      <table class="doc-table">${filasDatosOperacion}</table>

      <h2>2. Condiciones financieras</h2>
      <table class="doc-table">${filasCondiciones}</table>
      <p>${esc(t.condiciones)}</p>
      <ul>${t.detalleCondiciones.map((v) => `<li>${esc(v)}</li>`).join('')}</ul>

      ${filasBonificaciones ? `
        <h2>3. Bonificaciones aplicables</h2>
        <table class="doc-table doc-table--gastos">
          <thead><tr><th>Producto / Servicio</th><th>Bonificación</th></tr></thead>
          <tbody>${filasBonificaciones}</tbody>
        </table>
      ` : ''}

      <h2>4. Comisiones</h2>
      <table class="doc-table">
        ${row('Comisión de apertura', d.comisionApertura || '0 €')}
        ${row('Reembolso anticipado parcial (10 primeros años)', d.reembolsoParcial10 || '—')}
        ${row('Reembolso anticipado parcial (resto)', d.reembolsoParcialResto || '—')}
        ${row('Reembolso anticipado total (10 primeros años)', d.reembolsoTotal10 || '—')}
        ${row('Reembolso anticipado total (resto)', d.reembolsoTotalResto || '—')}
      </table>

      <h2>5. Desglose de gastos</h2>
      <table class="doc-table doc-table--gastos">
        <thead><tr><th>Concepto</th><th>Cliente</th><th>Entidad</th></tr></thead>
        <tbody>
          ${filasGastos}
          <tr>
            <th>Total</th>
            <td class="num"><strong>${esc(eur(d.totalCliente))}</strong></td>
            <td class="num"><strong>${esc(eur(d.totalEntidad))}</strong></td>
          </tr>
        </tbody>
      </table>

      <h2>6. Aportación y ahorros</h2>
      <ul>
        <li>Ahorros del cliente: ${esc(eur(d.ahorrosCliente))}.</li>
        <li>Arras / PYS: ${esc(eur(d.arrasPys))}.</li>
      </ul>

      <h2>7. Ahorro total a aportar por el cliente</h2>
      <div class="doc-highlight">
        <p>El cliente deberá aportar de fondos propios un total de:</p>
        <p class="big">${esc(eur(d.ahorroTotalAportar))}</p>
        <p><em>${esc(t.textoAhorro)}</em></p>
      </div>

      <h2>Conclusión</h2>
      <p>${esc(t.conclusion)}</p>

      <div class="doc-notas">${esc(t.notasLegales)}</div>
    `;
  }

  // ---------- Modales vista previa ----------
  function abrirModal() {
    const d = leerDatos();
    if (!d.nombreCliente) {
      alert('Introduce al menos el nombre del cliente.');
      form.elements.nombreCliente.focus();
      return;
    }
    ultimosDatos = d;
    previewContent.innerHTML = construirHTML(d);
    modal.hidden = false;
    document.body.style.overflow = 'hidden';
    modal.querySelector('.modal-body').scrollTop = 0;
  }
  function cerrarModal() {
    modal.hidden = true;
    document.body.style.overflow = '';
  }

  // =========================================================
  // ================  DOCX  =================================
  // =========================================================
  const BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' };

  function celdaDocx(texto, { bold = false, fill = null, width = null, align = null } = {}) {
    const parrafo = new Paragraph({
      alignment: align || AlignmentType.LEFT,
      children: [new TextRun({ text: String(texto ?? ''), bold, size: 22, font: FONT_DOCX })],
    });
    const opts = {
      children: [parrafo],
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
    };
    if (fill) opts.shading = { type: ShadingType.CLEAR, fill, color: 'auto' };
    if (width) opts.width = { size: width, type: WidthType.PERCENTAGE };
    return new TableCell(opts);
  }

  function tablaDocx(rows, { header = null, widths = null } = {}) {
    const trs = [];
    if (header) {
      trs.push(new TableRow({
        tableHeader: true,
        children: header.map((h, i) => celdaDocx(h, {
          bold: true, fill: '1E3A8A', width: widths?.[i] ?? null,
        })),
      }));
    }
    rows.forEach((r) => {
      trs.push(new TableRow({
        children: r.map((c, i) => {
          if (typeof c === 'object' && c !== null) {
            return celdaDocx(c.text, {
              bold: c.bold,
              fill: c.fill,
              width: widths?.[i] ?? null,
              align: c.align,
            });
          }
          return celdaDocx(c, { width: widths?.[i] ?? null });
        }),
      }));
    });
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: BORDER, bottom: BORDER, left: BORDER, right: BORDER,
        insideHorizontal: BORDER, insideVertical: BORDER,
      },
      rows: trs,
    });
  }

  function tablaDatosDocx(rows) {
    return tablaDocx(rows.map(([k, v]) => [
      { text: k, bold: true, fill: 'F1F5F9' },
      v,
    ]), { widths: [40, 60] });
  }

  function tituloDocx(texto) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 320, after: 160 },
      children: [new TextRun({ text: texto, bold: true, color: '1E3A8A', size: 26, font: FONT_DOCX })],
    });
  }

  function parrafoDocx(texto) {
    return new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ text: String(texto ?? ''), size: 22, font: FONT_DOCX })],
    });
  }

  function parrafoDestacadoDocx(texto, { size = 24, bold = false, color = null } = {}) {
    return new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({
        text: String(texto ?? ''), size, bold, font: FONT_DOCX,
        color: color || undefined,
      })],
    });
  }

  function vinetaDocx(texto) {
    return new Paragraph({
      bullet: { level: 0 },
      spacing: { after: 60 },
      children: [new TextRun({ text: String(texto ?? ''),
