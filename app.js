(() => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel } = window.docx;

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

  const STORAGE_KEY = 'propuesta_borrador_v2';
  const FONT_DOCX = 'Aptos, Calibri, Segoe UI, sans-serif';

  // Filas fijas de la tabla de gastos
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

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

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
      guardarBorrador();
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
      guardarBorrador();
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

  // ---------- Persistencia ----------
  function guardarBorrador() {
    const datos = {};

    // Campos simples
    for (const el of form.elements) {
      if (el.name && !el.name.endsWith('[]')) datos[el.name] = el.value;
    }

    // Bonificaciones
    datos._bonificaciones = [...listaBonificaciones.querySelectorAll('.fila-dinamica')].map((fila) => ({
      descripcion: fila.querySelector('input[name="bonif_desc[]"]')?.value || '',
      puntos: fila.querySelector('input[name="bonif_puntos[]"]')?.value || '',
    }));

    // Gastos
    datos._gastos = [...tbodyGastos.querySelectorAll('tr')].map((tr) => ({
      concepto: tr.querySelector('input[name="gasto_concepto[]"]')?.value || '',
      cliente: tr.querySelector('input[name="gasto_cliente[]"]')?.value || '',
      entidad: tr.querySelector('input[name="gasto_entidad[]"]')?.value || '',
    }));

    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(datos)); } catch (_) {}
  }

  function cargarBorrador() {
    try {
      const datos = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');

      // Campos simples
      for (const [k, v] of Object.entries(datos)) {
        if (k.startsWith('_')) continue;
        const el = form.elements[k];
        if (el) el.value = v;
      }

      // Bonificaciones
      listaBonificaciones.innerHTML = '';
      if (Array.isArray(datos._bonificaciones) && datos._bonificaciones.length) {
        datos._bonificaciones.forEach((b) => agregarBonificacion(b));
      } else {
        agregarBonificacion({ descripcion: 'Domiciliación de nómina', puntos: '0,50 p.p.' });
        agregarBonificacion({ descripcion: 'Seguro de hogar', puntos: '0,20 p.p.' });
      }

      // Gastos
      tbodyGastos.innerHTML = '';
      if (Array.isArray(datos._gastos) && datos._gastos.length) {
        datos._gastos.forEach((g) => agregarGasto(g.concepto, g.cliente, g.entidad));
      } else {
        inicializarGastos();
      }
    } catch (_) {
      inicializarGastos();
    }
  }

  function limpiarFormulario() {
    if (!confirm('¿Borrar todos los datos del formulario?')) return;
    form.reset();
    localStorage.removeItem(STORAGE_KEY);
    form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    listaBonificaciones.innerHTML = '';
    agregarBonificacion({ descripcion: 'Domiciliación de nómina', puntos: '0,50 p.p.' });
    agregarBonificacion({ descripcion: 'Seguro de hogar', puntos: '0,20 p.p.' });
    inicializarGastos();
    actualizarBloqueMixta();
  }

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
   
