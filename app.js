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
    })).filter((g) => g.concepto || g.cliente || g.entidad);
    return d;
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

  // ---------- Textos comunes ----------
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
      detalleCondiciones = [
        `Plazo: ${d.plazoAnos || '—'} años.`,
      ];
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

  // ---------- HTML de vista previa ----------
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
      children: [new TextRun({ text: String(texto ?? ''), size: 22, font: FONT_DOCX })],
    });
  }

  function construirDocumentoDOCX(d) {
    const t = textos(d);

    // Cabecera: datos de oficina/gestor
    const filasCabecera = [];
    if (d.oficina) filasCabecera.push(['Oficina', d.oficina]);
    if (d.gestor) filasCabecera.push(['Gestor', d.gestor]);
    if (d.expediente) filasCabecera.push(['Expediente', d.expediente]);
    if (d.emailGestor) filasCabecera.push(['Email gestor', d.emailGestor]);
    if (d.telefonoGestor) filasCabecera.push(['Teléfono gestor', d.telefonoGestor]);

    // Sección 1
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

    // Sección 2
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
    ];

    if (filasCabecera.length) {
      children.push(tablaDatosDocx(filasCabecera));
    }

    // 1
    children.push(tituloDocx('1. Datos de la operación'), tablaOperacion);

    // 2
    children.push(
      tituloDocx('2. Condiciones financieras'),
      tablaCondiciones,
      parrafoDocx(t.condiciones),
      ...t.detalleCondiciones.map(vinetaDocx),
    );

    // 3. Bonificaciones
    if (d._bonificaciones.length) {
      const tablaBonif = tablaDocx(
        d._bonificaciones.map((b) => [b.descripcion, b.puntos || '—']),
        { header: ['Producto / Servicio', 'Bonificación'], widths: [70, 30] }
      );
      children.push(tituloDocx('3. Bonificaciones aplicables'), tablaBonif);
    }

    // 4. Comisiones
    const tablaComisiones = tablaDatosDocx([
      ['Comisión de apertura', d.comisionApertura || '0 €'],
      ['Reembolso anticipado parcial (10 primeros años)', d.reembolsoParcial10 || '—'],
      ['Reembolso anticipado parcial (resto)', d.reembolsoParcialResto || '—'],
      ['Reembolso anticipado total (10 primeros años)', d.reembolsoTotal10 || '—'],
      ['Reembolso anticipado total (resto)', d.reembolsoTotalResto || '—'],
    ]);
    children.push(tituloDocx('4. Comisiones'), tablaComisiones);

    // 5. Desglose de gastos
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

    // 6. Aportación y ahorros
    children.push(
      tituloDocx('6. Aportación y ahorros'),
      vinetaDocx(`Ahorros del cliente: ${eur(d.ahorrosCliente)}.`),
      vinetaDocx(`Arras / PYS: ${eur(d.arrasPys)}.`),
    );

    // 7. Ahorro total
    children.push(
      tituloDocx('7. Ahorro total a aportar por el cliente'),
      parrafoDocx('El cliente deberá aportar de fondos propios un total de:'),
      parrafoDestacadoDocx(eur(d.ahorroTotalAportar), { size: 32, bold: true, color: '1E3A8A' }),
      parrafoDocx(t.textoAhorro),
    );

    // Conclusión
    children.push(
      tituloDocx('Conclusión'),
      parrafoDocx(t.conclusion),
    );

    // Notas legales
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

  // Botón "+ Añadir bonificación"
  document.querySelector('[data-add="bonificaciones"]').addEventListener('click', () => {
    agregarBonificacion();
    guardarBorrador();
  });

  // ---------- Init ----------
  document.addEventListener('DOMContentLoaded', () => {
    cargarBorrador();
    if (!form.elements.fechaDocumento.value) {
      form.elements.fechaDocumento.value = new Date().toISOString().split('T')[0];
    }
    actualizarBloqueMixta();
    if (!tbodyGastos.children.length) inicializarGastos();
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
