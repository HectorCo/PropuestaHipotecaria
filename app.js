(() => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel } = window.docx;
  const html2pdf = window.html2pdf;

  const form = document.getElementById('formPropuesta');
  const modal = document.getElementById('previewModal');
  const previewContent = document.getElementById('previewContent');
  const btnPreview = document.getElementById('btnPreview');
  const btnLimpiar = document.getElementById('btnLimpiar');
  const btnGenerar = document.getElementById('btnGenerar');
  const btnPDF = document.getElementById('btnPDF');
  const btnClose = document.getElementById('btnClose');

  const STORAGE_KEY = 'propuesta_borrador_v1';
  const FONT = 'Aptos, Calibri, Segoe UI, sans-serif';

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
  }

  function leerDatos() {
    const d = {};
    for (const el of form.elements) if (el.name) d[el.name] = (el.value || '').trim();
    return d;
  }

  // ---------- Textos comunes ----------
  function textos(d) {
    const conclusionDefault =
      `Esta propuesta equilibra una cuota mensual cómoda (el ${pct(d.ratioEndeudamiento)} de tus ingresos) ` +
      `con una financiación máxima. Es una oportunidad sólida para adquirir tu hogar con total seguridad financiera.`;
    return {
      intro:
        `Tras un análisis exhaustivo de vuestro perfil financiero, nos complace comunicaros que la operación ` +
        `ha sido calificada como ${d.calificacion || 'VIABLE'}. Esta propuesta destaca por ofrecer una financiación ` +
        `de (${pct(d.ltv)} LTV).`,
      condiciones:
        `Hemos seleccionado una modalidad ${d.tipoInteres || 'Fija'} que te protege frente a la volatilidad ` +
        `del mercado durante toda la vida de la hipoteca.`,
      conclusion: (d.textoConclusion || '').trim() || conclusionDefault,
    };
  }

  // ---------- HTML de vista previa ----------
  function construirHTML(d) {
    const t = textos(d);
    const row = (k, v) => `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`;
    const subtitle = [d.nombreCliente, d.fechaDocumento ? fmtFecha(d.fechaDocumento) : '']
      .filter(Boolean).join(' · ');

    return `
      <h1 class="doc-title">Propuesta de Financiación Hipotecaria</h1>
      <p class="doc-subtitle">${esc(subtitle)}</p>

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
        <li>Periodo (${esc(d.plazoAnos || '—')} años): Tipo de interés ${esc(d.tipoInteres || 'Fija')} al ${esc(pct(d.tin))} TIN.</li>
        <li>Vinculaciones: ${esc(d.vinculaciones || '—')}.</li>
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
    // Scroll arriba del todo
    modal.querySelector('.modal-body').scrollTop = 0;
  }

  function cerrarModal() {
    modal.hidden = true;
    document.body.style.overflow = '';
  }

  // ---------- DOCX ----------
  const BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' };

  function celda(texto, { bold = false, fill = null, width = null } = {}) {
    const opts = {
      children: [new Paragraph({
        children: [new TextRun({ text: texto, bold, size: 22, font: FONT })],
      })],
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
    };
    if (fill) opts.shading = { type: ShadingType.CLEAR, fill, color: 'auto' };
    if (width) opts.width = { size: width, type: WidthType.PERCENTAGE };
    return new TableCell(opts);
  }

  function tablaDatos(rows) {
    const tableRows = rows.map(([k, v]) => new TableRow({
      children: [
        celda(k, { bold: true, fill: 'F1F5F9', width: 40 }),
        celda(v, { width: 60 }),
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

  function titulo(texto) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 320, after: 160 },
      children: [new TextRun({ text: texto, bold: true, color: '1E3A8A', size: 26, font: FONT })],
    });
  }

  function parrafo(texto) {
    return new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ text: texto, size: 22, font: FONT })],
    });
  }

  function vineta(texto) {
    return new Paragraph({
      bullet: { level: 0 },
      spacing: { after: 60 },
      children: [new TextRun({ text: texto, size: 22, font: FONT })],
    });
  }

  function construirDocumento(d) {
    const t = textos(d);
    const subtitle = [d.nombreCliente, d.fechaDocumento ? fmtFecha(d.fechaDocumento) : '']
      .filter(Boolean).join(' · ');

    return new Document({
      creator: 'Generador de Propuestas',
      title: `Propuesta de financiación hipotecaria - ${d.nombreCliente || ''}`,
      // Fuente Aptos como estilo por defecto del documento
      styles: {
        default: {
          document: {
            run: { font: FONT, size: 22 },
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
              bold: true, size: 34, color: '1E3A8A', font: FONT,
            })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 320 },
            children: [new TextRun({
              text: subtitle, italics: true, color: '64748B', size: 20, font: FONT,
            })],
          }),

          titulo('1. Resumen operación'),
          parrafo(t.intro),
          tablaDatos([
            ['Capital Hipotecario', eur(d.capitalHipotecario)],
            ['Cuota Mensual Estimada', eur(d.cuotaMensual)],
            ['Ratio de Endeudamiento', pct(d.ratioEndeudamiento)],
            ['LTV (Financiación)', pct(d.ltv) + (d.precioVenta ? ' sobre el precio de venta' : '')],
          ]),

          titulo('2. Detalles de la Inversión'),
          parrafo('El presupuesto total de inversión se desglosa para garantizar total transparencia en cada paso de la compraventa.'),
          tablaDatos([
            ['Precio de venta', eur(d.precioVenta)],
            ['Valor de tasación objetivo', eur(d.valorTasacion)],
            ['Capital hipotecario', eur(d.capitalHipotecario)],
          ]),

          titulo('3. Condiciones de la Hipoteca'),
          parrafo(t.condiciones),
          vineta(`Periodo (${d.plazoAnos || '—'} años): Tipo de interés ${d.tipoInteres || 'Fija'} al ${pct(d.tin)} TIN.`),
          vineta(`Vinculaciones: ${d.vinculaciones || '—'}.`),

          titulo('4. Aportación y Ahorros'),
          parrafo(`Gracias a la estructura de financiación diseñada (LTV ${pct(d.ltv)}), la aportación de ahorros es la escogida, permitiéndote conservar capital para el futuro.`),
          vineta(`Ahorros del Cliente: ${eur(d.ahorrosCliente)}.`),
          vineta(`Arras / PYS: ${eur(d.arrasPys)}.`),
          vineta(`Valor de Tasación Objetivo confirmado: ${eur(d.valorTasacion)}.`),

          titulo('¿Por qué esta es vuestra mejor opción?'),
          parrafo(t.conclusion),
        ],
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
    return `Propuesta_${base}_${new Date().getFullYear()}.${ext}`;
  }

  async function descargarDOCX() {
    if (!ultimosDatos) return;
    const orig = btnGenerar.textContent;
    btnGenerar.disabled = true;
    btnGenerar.textContent = 'Generando…';
    try {
      const doc = construirDocumento(ultimosDatos);
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

  async function descargarPDF() {
    if (!ultimosDatos) return;
    const orig = btnPDF.textContent;
    btnPDF.disabled = true;
    btnPDF.textContent = 'Generando…';

    // Clonar la vista previa en un contenedor off-screen para capturarla limpia
    const clone = previewContent.cloneNode(true);
    clone.style.padding = '0';
    clone.style.width = 'auto';
    clone.style.maxWidth = 'none';
    clone.style.boxShadow = 'none';
    clone.style.borderRadius = '0';

    const wrapper = document.createElement('div');
    wrapper.style.cssText =
      'position:fixed;left:-10000px;top:0;width:186mm;background:#ffffff;';
    wrapper.appendChild(clone);
    document.body.appendChild(wrapper);

    try {
      await html2pdf().set({
        margin: [12, 12, 12, 12],              // mm
        filename: nombreArchivo('pdf'),
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: {
          scale: 2,
          useCORS: true,
          backgroundColor: '#ffffff',
          logging: false,
        },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: ['css', 'legacy'] },
      }).from(clone).save();
    } catch (err) {
      console.error(err);
      alert('Error al generar el PDF: ' + err.message);
    } finally {
      document.body.removeChild(wrapper);
      btnPDF.disabled = false;
      btnPDF.textContent = orig;
    }
  }

  // ---------- Eventos ----------
  form.addEventListener('input', guardarBorrador);
  form.addEventListener('change', guardarBorrador);
  btnPreview.addEventListener('click', abrirModal);
  btnLimpiar.addEventListener('click', limpiarFormulario);
  btnGenerar.addEventListener('click', descargarDOCX);
  btnPDF.addEventListener('click', descargarPDF);
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
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW error:', e));
    });
  }
})();
