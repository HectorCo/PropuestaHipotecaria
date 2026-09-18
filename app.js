(() => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel } = window.docx;

  const form = document.getElementById('formPropuesta');
  const btnGenerar = document.getElementById('btnGenerar');
  const btnLimpiar = document.getElementById('btnLimpiar');

  const STORAGE_KEY = 'propuesta_borrador_v1';

  // ---------- Utilidades de formato ----------
  const eur = (v) => {
    const s = String(v ?? '').trim();
    if (!s) return '—';
    // Si ya viene formateado con €, lo dejamos tal cual
    if (s.includes('€')) return s;
    // Si viene con coma decimal, la respetamos; si no, intentamos parsear
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

  // ---------- Construcción del DOCX ----------
  const BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' };

  function celda(texto, { bold = false, fill = null, width = null } = {}) {
    const opts = {
      children: [new Paragraph({
        children: [new TextRun({ text: texto, bold, size: 22 })],
      })],
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
    };
    if (fill) opts.shading = { type: ShadingType.CLEAR, fill, color: 'auto' };
    if (width) opts.width = { size: width, type: WidthType.PERCENTAGE };
    return new TableCell(opts);
  }

  function tablaDatos(rows, headers = null) {
    const tableRows = [];

    if (headers) {
      tableRows.push(new TableRow({
        tableHeader: true,
        children: headers.map((h, i) =>
          celda(h, { bold: true, fill: '1E3A8A', width: i === 0 ? 40 : 60 })
        ),
      }));
    }

    rows.forEach(([k, v]) => {
      tableRows.push(new TableRow({
        children: [
          celda(k, { bold: true, fill: 'F1F5F9', width: 40 }),
          celda(v, { width: 60 }),
        ],
      }));
    });

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
      children: [new TextRun({ text: texto, bold: true, color: '1E3A8A', size: 26 })],
    });
  }

  function parrafo(texto) {
    return new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ text: texto, size: 22 })],
    });
  }

  function vineta(texto) {
    return new Paragraph({
      bullet: { level: 0 },
      spacing: { after: 60 },
      children: [new TextRun({ text: texto, size: 22 })],
    });
  }

  function construirDocumento(d) {
    const conclusionDefault =
      `Esta propuesta equilibra una cuota mensual cómoda (el ${pct(d.ratioEndeudamiento)} de tus ingresos) ` +
      `con una financiación máxima. Es una oportunidad sólida para adquirir tu hogar con total seguridad financiera.`;

    const conclusion = (d.textoConclusion || '').trim() || conclusionDefault;

    const intro =
      `Tras un análisis exhaustivo de vuestro perfil financiero, nos complace comunicaros que la operación ` +
      `ha sido calificada como ${d.calificacion || 'VIABLE'}. Esta propuesta destaca por ofrecer una financiación ` +
      `de (${pct(d.ltv)} LTV).`;

    const condiciones =
      `Hemos seleccionado una modalidad ${d.tipoInteres || 'Fija'} que te protege frente a la volatilidad ` +
      `del mercado durante toda la vida de la hipoteca.`;

    return new Document({
      creator: 'Generador de Propuestas',
      title: `Propuesta de financiación hipotecaria - ${d.nombreCliente || ''}`,
      sections: [{
        properties: {
          page: { margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } },
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
            children: [new TextRun({ text: 'Propuesta de Financiación Hipotecaria', bold: true, size: 34, color: '1E3A8A' })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 320 },
            children: [new TextRun({
              text: `${d.nombreCliente || ''}${d.fechaDocumento ? ' · ' + fmtFecha(d.fechaDocumento) : ''}`,
              italics: true, color: '64748B', size: 20,
            })],
          }),

          // 1. Resumen
          titulo('1. Resumen operación'),
          parrafo(intro),
          tablaDatos([
            ['Capital Hipotecario', eur(d.capitalHipotecario)],
            ['Cuota Mensual Estimada', eur(d.cuotaMensual)],
            ['Ratio de Endeudamiento', pct(d.ratioEndeudamiento)],
            ['LTV (Financiación)', pct(d.ltv) + (d.precioVenta ? ' sobre el precio de venta' : '')],
          ]),

          // 2. Inversión
          titulo('2. Detalles de la Inversión'),
          parrafo('El presupuesto total de inversión se desglosa para garantizar total transparencia en cada paso de la compraventa.'),
          tablaDatos([
            ['Precio de venta', eur(d.precioVenta)],
            ['Valor de tasación objetivo', eur(d.valorTasacion)],
            ['Capital hipotecario', eur(d.capitalHipotecario)],
          ]),

          // 3. Condiciones
          titulo('3. Condiciones de la Hipoteca'),
          parrafo(condiciones),
          vineta(`Periodo (${d.plazoAnos || '—'} años): Tipo de interés ${d.tipoInteres || 'Fija'} al ${pct(d.tin)} TIN.`),
          vineta(`Vinculaciones: ${d.vinculaciones || '—'}.`),

          // 4. Aportación
          titulo('4. Aportación y Ahorros'),
          parrafo(`Gracias a la estructura de financiación diseñada (LTV ${pct(d.ltv)}), la aportación de ahorros es la escogida, permitiéndote conservar capital para el futuro.`),
          vineta(`Ahorros del Cliente: ${eur(d.ahorrosCliente)}.`),
          vineta(`Arras / PYS: ${eur(d.arrasPys)}.`),
          vineta(`Valor de Tasación Objetivo confirmado: ${eur(d.valorTasacion)}.`),

          // Conclusión
          titulo('¿Por qué esta es vuestra mejor opción?'),
          parrafo(conclusion),
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

  // ---------- Flujo principal ----------
  async function generar() {
    const d = {};
    for (const el of form.elements) if (el.name) d[el.name] = el.value.trim();

    if (!d.nombreCliente) {
      alert('Introduce al menos el nombre del cliente.');
      form.elements.nombreCliente.focus();
      return;
    }

    btnGenerar.disabled = true;
    const original = btnGenerar.textContent;
    btnGenerar.textContent = 'Generando…';

    try {
      const doc = construirDocumento(d);
      const blob = await Packer.toBlob(doc);
      const nombre = `Propuesta_${(d.nombreCliente || 'cliente').replace(/\s+/g, '_')}_${new Date().getFullYear()}.docx`;
      descargarBlob(blob, nombre);
    } catch (err) {
      console.error(err);
      alert('Error al generar el documento: ' + err.message);
    } finally {
      btnGenerar.disabled = false;
      btnGenerar.textContent = original;
    }
  }

  // ---------- Eventos ----------
  form.addEventListener('input', guardarBorrador);
  form.addEventListener('change', guardarBorrador);
  btnGenerar.addEventListener('click', generar);
  btnLimpiar.addEventListener('click', limpiarFormulario);

  // Init
  document.addEventListener('DOMContentLoaded', () => {
    cargarBorrador();
    if (!form.elements.fechaDocumento.value) {
      form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    }
  });

  // Registrar service worker
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW error:', e));
    });
  }
})();
