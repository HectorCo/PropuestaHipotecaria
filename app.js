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

  // Bloque condicional "Mixta"
  const bloqueMixta = document.getElementById('bloqueMixta');
  const selectTipoInteres = form.elements.tipoInteres;

  const STORAGE_KEY = 'propuesta_borrador_v1';
  const FONT_DOCX = 'Aptos, Calibri, Segoe UI, sans-serif';

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

  // ---------- Persistencia ----------
  function guardarBorrador() {
    const datos = {};
    for (const el of form.elements) if (el.name) datos[el.name] = el.value;
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(datos)); } catch (_) {}
  }

  function cargarBorrador() {
    try {
      const datos = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
      for (const [k, v] of Object.entries(datos)) {
        const el = form.elements[k];
        if (el) el.value = v;
      }
    } catch (_) {}
  }

  function limpiarFormulario() {
    if (!confirm('¿Borrar todos los datos del formulario?')) return;
    form.reset();
    localStorage.removeItem(STORAGE_KEY);
    form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    actualizarBloqueMixta();
  }

  function leerDatos() {
    const d = {};
    for (const el of form.elements) if (el.name) d[el.name] = (el.value || '').trim();
    return d;
  }

  // ---------- Visibilidad del bloque "Mixta" ----------
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

  // ---------- Textos comunes ----------
  function textos(d) {
    const conclusionDefault =
      `Esta propuesta equilibra una cuota mensual cómoda (el ${pct(d.ratioEndeudamiento)} de tus ingresos) ` +
      `con una financiación máxima. Es una oportunidad sólida para adquirir tu hogar con total seguridad financiera.`;

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
        `Hemos seleccionado una modalidad Mixta: un primer periodo a tipo fijo que te protege ` +
        `frente a las subidas de tipos, seguido de un periodo a tipo variable referenciado al ` +
        `${d.indiceVariable || 'Euríbor 12M'} más un diferencial.`;

      detalleCondiciones = [
        `Periodo fijo (${aniosFijos} años): TIN fijo al ${pct(d.tinFijo)}.`,
        `Periodo variable (${aniosVariables} años): ${d.indiceVariable || 'Euríbor 12M'} + ${pct(d.diferencialVariable)}.`,
        `Plazo total: ${d.plazoAnos || '—'} años.`,
        `Vinculaciones: ${d.vinculaciones || '—'}.`,
      ];
    } else {
      condiciones =
        `Hemos seleccionado una modalidad ${d.tipoInteres || 'Fija'} que te protege frente a la volatilidad ` +
        `del mercado durante toda la vida de la hipoteca.`;

      detalleCondiciones = [
        `Periodo (${d.plazoAnos || '—'} años): Tipo de interés ${d.tipoInteres || 'Fija'} al ${pct(d.tin)} TIN.`,
        `Vinculaciones: ${d.vinculaciones || '—'}.`,
      ];
    }

    return {
      intro:
        `Tras un análisis exhaustivo de vuestro perfil financiero, nos complace comunicaros que la operación ` +
        `ha sido calificada como ${d.calificacion || 'VIABLE'}. Esta propuesta destaca por ofrecer una financiación ` +
        `de (${pct(d.ltv)} LTV).`,
      condiciones,
      detalleCondiciones,
      conclusion: (d.textoConclusion || '').trim() || conclusionDefault,
      subtitle: [d.nombreCliente, d.fechaDocumento ? fmtFecha(d.fechaDocumento) : '']
        .filter(Boolean).join(' · '),
    };
  }

  // ---------- HTML de vista previa ----------
  function construirHTML(d) {
    const t = textos(d);
    const row = (k, v) => `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`;
    return `
      <h1 class="doc-title">Propuesta de Financiación Hipotecaria</h1>
      <p class="doc-subtitle">${esc(t.subtitle)}</p>

      <h2>1. Resumen operación</h2>
      <p>${esc(t.intro)}</p>
      <table class="doc-table">
        ${row('Capital Hipotecario', eur(d.capitalHipotecario))}
        ${row('Cuota Mensual Estimada', eur(d.cuotaMensual))}
        ${row('Ratio de Endeudamiento', pct(d.ratioEndeudamiento))}
        ${row('LTV (Financiación)', pct(d.ltv) + (d.precioVenta ? ' sobre el precio de venta' : ''))}
      </table>

      <h2>2. Detalles de la Inversión</h2>
      <p>El presupuesto total de inversión se desglosa para garantizar total transparencia en cada paso de la compraventa.</p>
      <table class="doc-table">
        ${row('Precio de venta', eur(d.precioVenta))}
        ${row('Valor de tasación objetivo', eur(d.valorTasacion))}
        ${row('Capital hipotecario', eur(d.capitalHipotecario))}
      </table>

      <h2>3. Condiciones de la Hipoteca</h2>
      <p>${esc(t.condiciones)}</p>
      <ul>
        ${t.detalleCondiciones.map((v) => `<li>${esc(v)}</li>`).join('')}
      </ul>

      <h2>4. Aportación y Ahorros</h2>
      <p>Gracias a la estructura de financiación diseñada (LTV ${esc(pct(d.ltv))}), la aportación de ahorros es la escogida, permitiéndote conservar capital para el futuro.</p>
      <ul>
        <li>Ahorros del Cliente: ${esc(eur(d.ahorrosCliente))}.</li>
        <li>Arras / PYS: ${esc(eur(d.arrasPys))}.</li>
        <li>Valor de Tasación Objetivo confirmado: ${esc(eur(d.valorTasacion))}.</li>
      </ul>

      <h2>¿Por qué esta es vuestra mejor opción?</h2>
      <p>${esc(t.conclusion)}</p>
    `;
  }

  // ---------- Modal ----------
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

  function celdaDocx(texto, { bold = false, fill = null, width = null } = {}) {
    const opts = {
      children: [new Paragraph({
        children: [new TextRun({ text: texto, bold, size: 22, font: FONT_DOCX })],
      })],
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
    };
    if (fill) opts.shading = { type: ShadingType.CLEAR, fill, color: 'auto' };
    if (width) opts.width = { size: width, type: WidthType.PERCENTAGE };
    return new TableCell(opts);
  }

  function tablaDatosDocx(rows) {
    const tableRows = rows.map(([k, v]) => new TableRow({
      children: [
        celdaDocx(k, { bold: true, fill: 'F1F5F9', width: 40 }),
        celdaDocx(v, { width: 60 }),
      ],
    }));
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: BORDER, bottom: BORDER, left: BORDER, right: BORDER,
        insideHorizontal: BORDER, insideVertical: BORDER,
      },
      rows: tableRows,
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
      children: [new TextRun({ text: texto, size: 22, font: FONT_DOCX })],
    });
  }

  function vinetaDocx(texto) {
    return new Paragraph({
      bullet: { level: 0 },
      spacing: { after: 60 },
      children: [new TextRun({ text: texto, size: 22, font: FONT_DOCX })],
    });
  }

  function construirDocumentoDOCX(d) {
    const t = textos(d);
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
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
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

          tituloDocx('1. Resumen operación'),
          parrafoDocx(t.intro),
          tablaDatosDocx([
            ['Capital Hipotecario', eur(d.capitalHipotecario)],
            ['Cuota Mensual Estimada', eur(d.cuotaMensual)],
            ['Ratio de Endeudamiento', pct(d.ratioEndeudamiento)],
            ['LTV (Financiación)', pct(d.ltv) + (d.precioVenta ? ' sobre el precio de venta' : '')],
          ]),

          tituloDocx('2. Detalles de la Inversión'),
          parrafoDocx('El presupuesto total de inversión se desglosa para garantizar total transparencia en cada paso de la compraventa.'),
          tablaDatosDocx([
            ['Precio de venta', eur(d.precioVenta)],
            ['Valor de tasación objetivo', eur(d.valorTasacion)],
            ['Capital hipotecario', eur(d.capitalHipotecario)],
          ]),

          tituloDocx('3. Condiciones de la Hipoteca'),
          parrafoDocx(t.condiciones),
          ...t.detalleCondiciones.map(vinetaDocx),

          tituloDocx('4. Aportación y Ahorros'),
          parrafoDocx(`Gracias a la estructura de financiación diseñada (LTV ${pct(d.ltv)}), la aportación de ahorros es la escogida, permitiéndote conservar capital para el futuro.`),
          vinetaDocx(`Ahorros del Cliente: ${eur(d.ahorrosCliente)}.`),
          vinetaDocx(`Arras / PYS: ${eur(d.arrasPys)}.`),
          vinetaDocx(`Valor de Tasación Objetivo confirmado: ${eur(d.valorTasacion)}.`),

          tituloDocx('¿Por qué esta es vuestra mejor opción?'),
          parrafoDocx(t.conclusion),
        ],
      }],
    });
  }

  // ---------- Descarga ----------
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
    return `Propuesta_${base}_${new Date().getFullYear()}.${ext}`;
  }

  async function descargarDOCX() {
    if (!ultimosDatos) return;
    const orig = btnGenerar.textContent;
    btnGenerar.disabled = true;
    btnGenerar.textContent = 'Generando…';
    try {
      const doc = construirDocumentoDOCX(ultimosDatos);
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
  form.addEventListener('input', guardarBorrador);
  form.addEventListener('change', guardarBorrador);
  selectTipoInteres.addEventListener('change', actualizarBloqueMixta);
  btnPreview.addEventListener('click', abrirModal);
  btnLimpiar.addEventListener('click', limpiarFormulario);
  btnGenerar.addEventListener('click', descargarDOCX);
  btnClose.addEventListener('click', cerrarModal);
  modal.addEventListener('click', (e) => {
    if (e.target.matches('[data-close]')) cerrarModal();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !modal.hidden) cerrarModal();
  });

  // ---------- Init ----------
  document.addEventListener('DOMContentLoaded', () => {
    cargarBorrador();
    if (!form.elements.fechaDocumento.value) {
      form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    }
    actualizarBloqueMixta();
  });

  // Service Worker con auto-actualización
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', async () => {
      try {
        const reg = await navigator.serviceWorker.register('sw.js');
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
      } catch (e) {
        console.warn('SW error:', e);
      }
    });
  }
})();
