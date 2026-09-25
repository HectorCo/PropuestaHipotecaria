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
  function cargarContadores() {
    try { return JSON.parse(localStorage.getItem(COUNTER_KEY) || '{}'); }
    catch { return {}; }
  }
  function guardarContadores(c) {
    try { localStorage.setItem(COUNTER_KEY, JSON.stringify(c)); } catch (_) {}
  }

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

  function mayorSecuenciaEnBorradores(clave) {
    const drafts = getDrafts();
    let max = 0;
    for (const d of drafts) {
      const exp = (d?.datos?.expediente || '').trim();
      if (!exp || exp.length < 6) continue;
      if (exp.slice(0, 6) !== clave) continue;
      const sec = parseInt(exp.slice(6), 10);
      if (!isNaN(sec) && sec > max) max = sec;
    }
    return max;
  }

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

  function asegurarExpediente() {
    const input = form.elements.expediente;
    if (!input) return '';
    const actual = (input.value || '').trim();
    if (actual) return actual;
    const fecha = form.elements.fechaDocumento?.value || '';
    const nuevo = generarExpediente(fecha);
    input.value = nuevo;
    return nuevo;
  }

  // =========================================================
  // ============  CÁLCULOS AUTOMÁTICOS  =====================
  // =========================================================

  // Parsea un valor monetario aceptando formatos español (1.234,56),
  // inglés (1234.56) o número plano (1234).
  // Devuelve un número o 0 si no se puede parsear.
  function parseNumero(v) {
    const s = String(v ?? '').trim();
    if (!s) return 0;
    // Quitamos símbolos de moneda y espacios
    const limpio = s.replace(/[€$\s]/g, '');
    if (!limpio) return 0;
    // Si contiene coma y punto, la coma es decimal (formato español)
    if (limpio.includes(',') && limpio.includes('.')) {
      const n = parseFloat(limpio.replace(/\./g, '').replace(',', '.'));
      return isNaN(n) ? 0 : n;
    }
    // Si solo contiene coma, la coma es decimal
    if (limpio.includes(',')) {
      const n = parseFloat(limpio.replace(',', '.'));
      return isNaN(n) ? 0 : n;
    }
    // Sin coma: parseo directo (los puntos pueden ser miles o decimales,
    // en cualquiera de los dos casos parseFloat funciona)
    const n = parseFloat(limpio);
    return isNaN(n) ? 0 : n;
  }

  // Formatea un número como moneda española sin símbolo
  function formatoMoneda(n) {
    return new Intl.NumberFormat('es-ES', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(n || 0);
  }

  // Suma los valores de una columna de la tabla de gastos
  function sumarColumnaGastos(columna) {
    let total = 0;
    tbodyGastos.querySelectorAll('tr').forEach((tr) => {
      const input = tr.querySelector(`input[name="gasto_${columna}[]"]`);
      if (input) total += parseNumero(input.value);
    });
    return total;
  }

  // Recalcula todos los campos dependientes (totales de gastos y ahorro total)
  function recalcularTotales() {
    // 1) Totales de gastos
    const totalCliente = sumarColumnaGastos('cliente');
    const totalEntidad = sumarColumnaGastos('entidad');

    const inputTotalCliente = form.elements.totalCliente;
    const inputTotalEntidad = form.elements.totalEntidad;
    if (inputTotalCliente) inputTotalCliente.value = formatoMoneda(totalCliente);
    if (inputTotalEntidad) inputTotalEntidad.value = formatoMoneda(totalEntidad);

    // 2) Ahorro total a aportar:
    //    Precio de compra − Importe financiado + Total cliente
    const precioCompra = parseNumero(form.elements.precioCompra?.value);
    const importeFinanciado = parseNumero(form.elements.capitalHipotecario?.value);
    const ahorroTotal = precioCompra - importeFinanciado + totalCliente;

    const inputAhorro = form.elements.ahorroTotalAportar;
    if (inputAhorro) inputAhorro.value = formatoMoneda(ahorroTotal);
  }

  // ---------- Utilidades ----------
  const eur = (v) => {
    const s = String(v ?? '').trim();
    if (!s) return '—';
    if (s.includes('€')) return s;
    const n = parseNumero(s);
    if (!n) return s.endsWith('€') ? s : (s + ' €');
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
      recalcularTotales();
    });
    // Recalcula al escribir en las celdas Cliente/Entidad
    tr.querySelectorAll('input[name="gasto_cliente[]"], input[name="gasto_entidad[]"]').forEach((input) => {
      input.addEventListener('input', recalcularTotales);
    });
    return tr;
  }

  function agregarGasto(concepto, cliente, entidad) {
    tbodyGastos.appendChild(crearFilaGasto(concepto, cliente, entidad));
  }

  function inicializarGastos() {
    tbodyGastos.innerHTML = '';
    FILAS_GASTOS_DEFAULT.forEach((c) => agregarGasto(c, '', ''));
    recalcularTotales();
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
   
