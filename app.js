(() => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel,
          ImageRun } = window.docx;

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
  const COUNTER_KEY = 'propuesta_expediente_counter_v1';
  const FONT_DOCX = 'Aptos, Calibri, Segoe UI, sans-serif';

  const FIXED_FIELDS = {
    gestor: 'Héctor Company',
    emailGestor: 'bgg@bgestionglobal.net',
    telefonoGestor: '602250255',
  };

  const FILAS_GASTOS_DEFAULT = [
    'Registro de la propiedad',
    'Notaría',
    'Gestoría',
    'Impuestos',
    'Tasación',
    'Nota simple',
    'Seguro de daños',
  ];

  const LOGO_IZQ_DEFAULT = 'logo-b.png';
  const LOGO_DER_DEFAULT = 'logo-blanco.png';
  const LOGO_MAX_WIDTH_PX = 150;
  const LOGO_MAX_HEIGHT_PX = 70;

  let ultimosDatos = null;
  let currentDraftId = null;
  let currentDraftName = '';
  let logoIzquierdoDataUrl = null;
  let logoDerechoDataUrl = null;

  // =========================================================
  // ============  EXPEDIENTE AUTOMÁTICO  ====================
  // =========================================================
  // Formato: AAAAMM#####  (año 4 dígitos + mes 2 dígitos + secuencia 5 dígitos)
  // Ejemplo: 20260900001  -> primera propuesta de septiembre de 2026
  //          20260900002  -> segunda propuesta del mismo mes
  //          20261000001  -> primera propuesta de octubre de 2026

  function cargarContadores() {
    try { return JSON.parse(localStorage.getItem(COUNTER_KEY) || '{}'); }
    catch { return {}; }
  }
  function guardarContadores(c) {
    try { localStorage.setItem(COUNTER_KEY, JSON.stringify(c)); } catch (_) {}
  }

  // Devuelve "YYYYMM" a partir de una fecha ISO (yyyy-mm-dd) o de hoy
  function claveMes(fechaIso) {
    let d;
    if (fechaIso) {
      d = new Date(fechaIso + 'T00:00:00');
      if (isNaN(d.getTime())) d = new Date();
    } else {
      d = new Date();
    }
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    return `${y}${m}`;
  }

  // Recorre los borradores y devuelve el mayor número de secuencia
  // encontrado para un mes dado (según el campo `expediente`).
  function mayorSecuenciaEnBorradores(clave) {
    const drafts = getDrafts();
    let max = 0;
    const prefijo = clave; // 6 dígitos
    for (const d of drafts) {
      const exp = (d?.datos?.expediente || '').trim();
      if (!exp || exp.length < 6) continue;
      if (exp.slice(0, 6) !== prefijo) continue;
      const sec = parseInt(exp.slice(6), 10);
      if (!isNaN(sec) && sec > max) max = sec;
    }
    return max;
  }

  // Reserva un nuevo número de expediente para el mes indicado.
  // Se apoya en el contador de localStorage pero también tiene en cuenta
  // los borradores existentes (por si vienen importados de otro equipo).
  function generarExpediente(fechaIso) {
    const clave = claveMes(fechaIso);
    const contadores = cargarContadores();
    const maxBorradores = mayorSecuenciaEnBorradores(clave);
    const actual = contadores[clave] || 0;
    const base = Math.max(actual, maxBorradores);
    const siguiente = base + 1;
    contadores[clave] = siguiente;
    guardarContadores(contadores);
    return `${clave}${String(siguiente).padStart(5, '0')}`;
  }

  // Asigna un expediente al borrador activo si aún no tiene uno.
  // Devuelve el expediente resultante.
  function asegurarExpediente() {
    const input = form.elements.expediente;
    if (!input) return '';
    const actual = (input.value || '').trim();
    if (actual) return actual; // ya tiene uno, se respeta
    const fecha = form.elements.fechaDocumento?.value || '';
    const nuevo = generarExpediente(fecha);
    input.value = nuevo;
    return nuevo;
  }

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

  // Convierte un dataURL a Uint8Array
  function dataUrlToUint8(dataUrl) {
    const base64 = dataUrl.split(',')[1];
    const bin = atob(base64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function dataUrlMime(dataUrl) {
    const m = /^data:([^;]+);/.exec(dataUrl || '');
    return m ? m[1] : 'image/png';
  }

  function docxImageType(mime) {
    if (mime.includes('png')) return 'png';
    if (mime.includes('jpeg') || mime.includes('jpg')) return 'jpg';
    if (mime.includes('gif')) return 'gif';
    if (mime.includes('bmp')) return 'bmp';
    return 'png';
  }

  function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async function urlToDataUrl(url) {
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error('No se pudo cargar ' + url);
    const blob = await res.blob();
    return await fileToDataUrl(blob);
  }

  async function resizeDataUrl(dataUrl, maxW, maxH) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        let w = img.naturalWidth;
        let h = img.naturalHeight;
        const ratio = Math.min(maxW / w, maxH / h, 1);
        w = Math.round(w * ratio);
        h = Math.round(h * ratio);
        resolve({ width: w, height: h, dataUrl });
      };
      img.onerror = () => resolve({ width: maxW, height: maxH, dataUrl });
      img.src = dataUrl;
    });
  }

  function aplicarCamposFijos() {
    for (const [k, v] of Object.entries(FIXED_FIELDS)) {
      const el = form.elements[k];
      if (el) el.value = v;
    }
  }

  // ---------- Filas dinámicas ----------
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

  // ---------- Lectura / escritura ----------
  function leerDatos() {
    const d = {};
    for (const el of form.elements) {
      if (el.name && !el.name.endsWith('[]') && el.type !== 'file') {
        d[el.name] = (el.value || '').trim();
      }
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
    if (logoIzquierdoDataUrl) d._logoIzq = logoIzquierdoDataUrl;
    if (logoDerechoDataUrl) d._logoDer = logoDerechoDataUrl;
    return d;
  }

  function aplicarDatosAlFormulario(datos) {
    if (!datos) datos = {};
    form.reset();
    aplicarCamposFijos();

    for (const [k, v] of Object.entries(datos)) {
      if (k.startsWith('_')) continue;
      if (k in FIXED_FIELDS) continue;
      const el = form.elements[k];
      if (el) el.value = v ?? '';
    }

    logoIzquierdoDataUrl = datos._logoIzq || null;
    logoDerechoDataUrl = datos._logoDer || null;

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
  function findDraft(id) { return getDrafts().find((d) => d.id === id); }
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

  function guardarBorradorActual() {
    // Antes de guardar, asegura que hay expediente
    asegurarExpediente();
    if (!currentDraftId) { guardarComoNuevoBorrador(); return; }
    const datos = leerDatos();
    const draft = findDraft(currentDraftId);
    if (!draft) { guardarComoNuevoBorrador(); return; }
    draft.datos = datos;
    draft.updatedAt = new Date().toISOString();
    if (!draft.nombre && datos.nombreCliente) draft.nombre = datos.nombreCliente;
    upsertDraft(draft);
    currentDraftName = draft.nombre;
    actualizarBarra();
  }

  function guardarComoNuevoBorrador() {
    asegurarExpediente();
    const datos = leerDatos();
    const sugerido = datos.nombreCliente || 'Nuevo borrador';
    const nombre = prompt('Nombre del borrador:', sugerido);
    if (nombre === null) return;
    const nombreFinal = (nombre || '').trim() || 'Sin nombre';
    const id = generarId();
    const now = new Date().toISOString();
    const draft = { id, nombre: nombreFinal, createdAt: now, updatedAt: now, datos };
    upsertDraft(draft);
    currentDraftId = id;
    currentDraftName = nombreFinal;
    actualizarBarra();
  }

  function nuevoBorrador() {
    if (!confirm('¿Crear un borrador nuevo? Los datos no guardados se perderán.')) return;
    currentDraftId = null;
    currentDraftName = '';
    logoIzquierdoDataUrl = null;
    logoDerechoDataUrl = null;
    localStorage.removeItem(STORAGE_KEY);
    form.reset();
    form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    aplicarCamposFijos();
    // El expediente se generará al guardar o hacer vista previa.
    form.elements.expediente.value = '';
    inicializarBonificaciones();
    inicializarGastos();
    actualizarBloqueMixta();
    actualizarBarra();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

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
    if (currentDraftId === id) { currentDraftName = nombreFinal; actualizarBarra(); }
    renderizarListaDrafts();
  }

  function duplicarBorrador(id) {
    const draft = findDraft(id);
    if (!draft) return;
    const copia = JSON.parse(JSON.stringify(draft));
    copia.id = generarId();
    copia.nombre = draft.nombre + ' (copia)';
    copia.createdAt = new Date().toISOString();
    copia.updatedAt = new Date().toISOString();
    // El duplicado conserva el expediente original: es una copia del mismo trabajo.
    upsertDraft(copia);
    renderizarListaDrafts();
  }

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

  function importarBorradores(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        const lista = Array.isArray(data) ? data : [data];
        const drafts = getDrafts();
        let añadidos = 0;
        // 1) Primera pasada: añade todos los borradores importados
        lista.forEach((item) => {
          if (!item || typeof item !== 'object' || !item.datos) return;
          const copia = JSON.parse(JSON.stringify(item));
          copia.id = generarId();
          copia.nombre = (copia.nombre || 'Importado').trim() || 'Importado';
          copia.createdAt = copia.createdAt || new Date().toISOString();
          copia.updatedAt = new Date().toISOString();
          copia.importedAt = new Date().toISOString();
          for (const [k, v] of Object.entries(FIXED_FIELDS)) copia.datos[k] = v;
          drafts.push(copia);
          añadidos++;
        });
        saveDrafts(drafts);
        // 2) Segunda pasada: recalcula los contadores por mes para que el
        //    próximo expediente que se genere continúe la serie correcta.
        recalcularContadoresDesdeBorradores();
        renderizarListaDrafts();
        alert(añadidos === 0
          ? 'No se encontraron borradores válidos en el archivo.'
          : `${añadidos} borrador(es) importado(s) correctamente.`);
      } catch (err) {
        console.error(err);
        alert('Error al importar el archivo: ' + err.message);
      }
    };
    reader.onerror = () => alert('Error al leer el archivo.');
    reader.readAsText(file);
  }

  // Recalcula el contador por mes a partir de los expedientes existentes
  // en los borradores. Se usa al importar para mantener la serie correcta.
  function recalcularContadoresDesdeBorradores() {
    const drafts = getDrafts();
    const contadores = cargarContadores();
    for (const d of drafts) {
      const exp = (d?.datos?.expediente || '').trim();
      if (!exp || exp.length < 7) continue;
      const clave = exp.slice(0, 6);
      const sec = parseInt(exp.slice(6), 10);
      if (isNaN(sec)) continue;
      if ((contadores[clave] || 0) < sec) contadores[clave] = sec;
    }
    guardarContadores(contadores);
  }

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
            ${d.datos?.expediente ? 'Exp. ' + esc(d.datos.expediente) + ' · ' : ''}
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
  // ================  TEXTOS  ===============================
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

    const srcIzq = d._logoIzq || LOGO_IZQ_DEFAULT;
    const srcDer = d._logoDer || LOGO_DER_DEFAULT;

    return `
      <div class="doc-header-logos">
        <img class="logo-left" src="${esc(srcIzq)}" alt="Logo izquierdo" />
        <img class="logo-right" src="${esc(srcDer)}" alt="Logo derecho" />
      </div>

      <h1 class="doc-title">Propuesta de Financiación Hipotecaria</h1>
      <p class="doc-subtitle">${esc(t.subtitle)}</p>

      <table class="doc-table">
        ${d.oficina ? row('Oficina', d.oficina) : ''}
        ${d.expediente ? row('Expediente', d.expediente) : ''}
        ${d.gestor ? row('Gestor', d.gestor) : ''}
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

  function abrirModal() {
    // Al abrir la vista previa, asegura que hay expediente
    asegurarExpediente();
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
              bold: c.bold, fill: c.fill,
              width: widths?.[i] ?? null, align: c.align,
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

  function celdaLogoDocx(runOrText, { align = AlignmentType.LEFT } = {}) {
    return new TableCell({
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      },
      width: { size: 50, type: WidthType.PERCENTAGE },
      children: [new Paragraph({ alignment: align, children: [runOrText] })],
    });
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
      children: [new TextRun({ text: String(texto ?? ''), size: 22, font: FONT_DOCX })],
    });
  }

  async function construirImageRun(dataUrl) {
    const resized = await resizeDataUrl(dataUrl, LOGO_MAX_WIDTH_PX, LOGO_MAX_HEIGHT_PX);
    const bytes = dataUrlToUint8(resized.dataUrl);
    const mime = dataUrlMime(resized.dataUrl);
    return new ImageRun({
      data: bytes,
      transformation: { width: resized.width, height: resized.height },
      type: docxImageType(mime),
    });
  }

  async function obtenerLogoDataUrl(subido, urlDefault) {
    if (subido) return subido;
    try { return await urlToDataUrl(urlDefault); }
    catch (_) { return null; }
  }

  async function construirDocumentoDOCX(d) {
    const t = textos(d);

    const izqDataUrl = await obtenerLogoDataUrl(d._logoIzq, LOGO_IZQ_DEFAULT);
    const derDataUrl = await obtenerLogoDataUrl(d._logoDer, LOGO_DER_DEFAULT);
    const izqRun = izqDataUrl ? await construirImageRun(izqDataUrl) : new TextRun({ text: '' });
    const derRun = derDataUrl ? await construirImageRun(derDataUrl) : new TextRun({ text: '' });

    const tablaLogos = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      },
      rows: [
        new TableRow({
          children: [
            celdaLogoDocx(izqRun, { align: AlignmentType.LEFT }),
            celdaLogoDocx(derRun, { align: AlignmentType.RIGHT }),
          ],
        }),
      ],
    });

    const filasCabecera = [];
    if (d.oficina) filasCabecera.push(['Oficina', d.oficina]);
    if (d.expediente) filasCabecera.push(['Expediente', d.expediente]);
    if (d.gestor) filasCabecera.push(['Gestor', d.gestor]);
    if (d.emailGestor) filasCabecera.push(['Email gestor', d.emailGestor]);
    if (d.telefonoGestor) filasCabecera.push(['Teléfono gestor', d.telefonoGestor]);

    const tablaOperacion = tablaDatosDocx([
      ['Precio de compra', eur(d.precioCompra)],
      ['Importe del préstamo hipotecario', eur(d.capitalHipotecario)],
      ['Plazo', (d.plazoAnos || '—') + ' años'],
      ['Tipo elegido', d.tipoInteres || '—'],
      ['Producto comercial', d.productoComercial || '—'],
      ['Finalidad', d.finalidad || '—'],
      ['Tipo de inmueble', d.tipoInmueble || '—'],
      ['Ubicación del inmueble', d.ubicacionInmueble || '—'],
    ]);

    const filasCond = [
      ['Cuota mensual — con bonificación', eur(d.cuotaBonificada)],
      ['Cuota mensual — sin bonificación', eur(d.cuotaSinBonificar)],
      ['TIN — con bonificación', pct(d.tinBonificado)],
      ['TIN — sin bonificación', pct(d.tinSinBonificar)],
      ['TAE — con bonificación', pct(d.taeBonificada)],
      ['TAE — sin bonificación', pct(d.taeSinBonificar)],
      ['Importe total adeudado — con bonificación', eur(d.importeTotalBonificado)],
      ['Importe total adeudado — sin bonificación', eur(d.importeTotalSinBonificar)],
    ];
    if (d.numeroCuotas) filasCond.push(['Número de cuotas', d.numeroCuotas]);
    const tablaCondiciones = tablaDatosDocx(filasCond);

    const children = [
      tablaLogos,
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 80 },
        children: [new TextRun({
          text: 'Propuesta de Financiación Hipotecaria',
          bold: true, size: 34, color: '1E3A8A', font: FONT_DOCX,
        })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 320 },
        children: [new TextRun({
          text: t.subtitle, italics: true, color: '64748B', size: 20, font: FONT_DOCX,
        })],
      }),
    ];

    if (filasCabecera.length) children.push(tablaDatosDocx(filasCabecera));

    children.push(tituloDocx('1. Datos de la operación'), tablaOperacion);

    children.push(
      tituloDocx('2. Condiciones financieras'),
      tablaCondiciones,
      parrafoDocx(t.condiciones),
      ...t.detalleCondiciones.map(vinetaDocx),
    );

    if (d._bonificaciones.length) {
      const tablaBonif = tablaDocx(
        d._bonificaciones.map((b) => [b.descripcion, b.puntos || '—']),
        { header: ['Producto / Servicio', 'Bonificación'], widths: [70, 30] }
      );
      children.push(tituloDocx('3. Bonificaciones aplicables'), tablaBonif);
    }

    const tablaComisiones = tablaDatosDocx([
      ['Comisión de apertura', d.comisionApertura || '0 €'],
      ['Reembolso anticipado parcial (10 primeros años)', d.reembolsoParcial10 || '—'],
      ['Reembolso anticipado parcial (resto)', d.reembolsoParcialResto || '—'],
      ['Reembolso anticipado total (10 primeros años)', d.reembolsoTotal10 || '—'],
      ['Reembolso anticipado total (resto)', d.reembolsoTotalResto || '—'],
    ]);
    children.push(tituloDocx('4. Comisiones'), tablaComisiones);

    const filasGastosDocx = d._gastos.map((g) => [
      g.concepto,
      g.cliente ? eur(g.cliente) : '—',
      g.entidad ? eur(g.entidad) : '—',
    ]);
    filasGastosDocx.push([
      { text: 'Total', bold: true },
      { text: eur(d.totalCliente), bold: true, align: AlignmentType.RIGHT },
      { text: eur(d.totalEntidad), bold: true, align: AlignmentType.RIGHT },
    ]);
    const tablaGastos = tablaDocx(filasGastosDocx, {
      header: ['Concepto', 'Cliente', 'Entidad'],
      widths: [50, 25, 25],
    });
    children.push(tituloDocx('5. Desglose de gastos'), tablaGastos);

    children.push(
      tituloDocx('6. Aportación y ahorros'),
      vinetaDocx(`Ahorros del cliente: ${eur(d.ahorrosCliente)}.`),
      vinetaDocx(`Arras / PYS: ${eur(d.arrasPys)}.`),
    );

    children.push(
      tituloDocx('7. Ahorro total a aportar por el cliente'),
      parrafoDocx('El cliente deberá aportar de fondos propios un total de:'),
      parrafoDestacadoDocx(eur(d.ahorroTotalAportar), { size: 32, bold: true, color: '1E3A8A' }),
      parrafoDocx(t.textoAhorro),
    );

    children.push(tituloDocx('Conclusión'), parrafoDocx(t.conclusion));

    children.push(
      new Paragraph({
        spacing: { before: 400, after: 120 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1', space: 8 } },
        children: [new TextRun({
          text: t.notasLegales,
          size: 18, italics: true, color: '64748B', font: FONT_DOCX,
        })],
      }),
    );

    return new Document({
      creator: 'Generador de Propuestas',
      title: `Propuesta de financiación hipotecaria - ${d.nombreCliente || ''}`,
      styles: {
        default: {
          document: {
            run: { font: FONT_DOCX, size: 22 },
            paragraph: { spacing: { after: 120 } },
          },
        },
      },
      sections: [{
        properties: {
          page: { margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } },
        },
        children,
      }],
    });
  }

  // ---------- Descargas ----------
  function descargarBlob(blob, nombre) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nombre;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function nombreArchivo(ext) {
    const base = (ultimosDatos?.nombreCliente || 'cliente').replace(/\s+/g, '_');
    const exp = (ultimosDatos?.expediente || '').trim();
    return exp
      ? `Propuesta_${exp}_${base}.${ext}`
      : `Propuesta_${base}_${new Date().getFullYear()}.${ext}`;
  }

  async function descargarDOCX() {
    if (!ultimosDatos) return;
    const orig = btnGenerar.textContent;
    btnGenerar.disabled = true;
    btnGenerar.textContent = 'Generando…';
    try {
      const doc = await construirDocumentoDOCX(ultimosDatos);
      const blob = await Packer.toBlob(doc);
      descargarBlob(blob, nombreArchivo('docx'));
    } catch (err) {
      console.error(err);
      alert('Error al generar el DOCX: ' + err.message);
    } finally {
      btnGenerar.disabled = false;
      btnGenerar.textContent = orig;
    }
  }

  // ---------- Eventos ----------
  form.addEventListener('input', autosave);
  form.addEventListener('change', autosave);
  selectTipoInteres.addEventListener('change', () => { actualizarBloqueMixta(); autosave(); });

  form.elements.logoIzquierdoFile?.addEventListener('change', async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      logoIzquierdoDataUrl = await fileToDataUrl(f);
      autosave();
    } catch (err) { alert('No se pudo cargar el logo: ' + err.message); }
    e.target.value = '';
  });

  form.elements.logoDerechoFile?.addEventListener('change', async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      logoDerechoDataUrl = await fileToDataUrl(f);
      autosave();
    } catch (err) { alert('No se pudo cargar el logo: ' + err.message); }
    e.target.value = '';
  });

  btnPreview.addEventListener('click', abrirModal);
  btnGenerar.addEventListener('click', descargarDOCX);
  btnClose.addEventListener('click', cerrarModal);
  modal.addEventListener('click', (e) => { if (e.target.matches('[data-close]')) cerrarModal(); });

  btnNuevo.addEventListener('click', nuevoBorrador);
  btnGuardar.addEventListener('click', guardarBorradorActual);
  btnGuardarComo.addEventListener('click', guardarComoNuevoBorrador);
  btnMisBorradores.addEventListener('click', abrirDraftsModal);

  draftsModal.addEventListener('click', (e) => {
    if (e.target.matches('[data-close-drafts]')) cerrarDraftsModal();
  });
  btnExportarTodos.addEventListener('click', exportarTodos);
  btnImportar.addEventListener('click', () => inputImportar.click());
  inputImportar.addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (file) importarBorradores(file);
    inputImportar.value = '';
  });

  btnLimpiar.addEventListener('click', () => {
    if (!confirm('¿Vaciar todos los campos del formulario? El borrador activo no se eliminará.')) return;
    form.reset();
    form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    logoIzquierdoDataUrl = null;
    logoDerechoDataUrl = null;
    aplicarCamposFijos();
    form.elements.expediente.value = '';
    inicializarBonificaciones();
    inicializarGastos();
    actualizarBloqueMixta();
    autosave();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    if (!modal.hidden) cerrarModal();
    if (!draftsModal.hidden) cerrarDraftsModal();
  });

  document.querySelector('[data-add="bonificaciones"]').addEventListener('click', () => {
    agregarBonificacion();
    autosave();
  });

  // ---------- Init ----------
  document.addEventListener('DOMContentLoaded', () => {
    cargarScratch();
    if (!form.elements.fechaDocumento.value) {
      form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    }
    aplicarCamposFijos();
    actualizarBloqueMixta();
    actualizarBarra();
  });

  // =========================================================
  // ==========  SERVICE WORKER (sw-v3.js)  ==================
  // =========================================================
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', async () => {
      try {
        const regs = await navigator.serviceWorker.getRegistrations();
        for (const r of regs) {
          const url = (r.active || r.installing || r.waiting || {}).scriptURL || '';
          if (url.indexOf('sw-v3.js') === -1) {
            try { await r.unregister(); } catch (_) {}
          }
        }
        const reg = await navigator.serviceWorker.register('sw-v3.js');
        reg.update();
        if (reg.waiting) reg.waiting.postMessage({ type: 'SKIP_WAITING' });
        reg.addEventListener('updatefound', () => {
          const nuevo = reg.installing;
          if (!nuevo) return;
          nuevo.addEventListener('statechange', () => {
            if (nuevo.state === 'installed' && navigator.serviceWorker.controller) {
              window.location.reload();
            }
          });
        });
        setInterval(() => {
          navigator.serviceWorker.getRegistration().then((r) => r?.update());
        }, 60 * 1000);
        window.addEventListener('focus', () => {
          navigator.serviceWorker.getRegistration().then((r) => r?.update());
        });
      } catch (e) {
        console.warn('SW error:', e);
      }
    });
  }
})();
