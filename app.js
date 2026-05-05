/* ============================================================
   AGROPECUARIA EL CHURQUI — app.js v7.0
   Correcciones:
   1. Panel principal: solo productos vigentes con stock > 0
   2. Salidas: mermas/pérdidas con motivo propio
   3. Fecha de fabricación + límite de 2 años para vencimiento
   4. Excel mejorado con formato profesional igual al PDF
   ============================================================ */

// ──────────────────────────────────────────────
// BASE DE DATOS — IndexedDB
// ──────────────────────────────────────────────
const DB_NAME    = 'churqui_db';
const DB_VERSION = 2;
let db;

function abrirDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const d = e.target.result;
      if (!d.objectStoreNames.contains('productos'))
        d.createObjectStore('productos', { keyPath: 'id', autoIncrement: true });
      if (!d.objectStoreNames.contains('entradas')) {
        const s = d.createObjectStore('entradas', { keyPath: 'id', autoIncrement: true });
        s.createIndex('fecha_hora', 'fecha_hora');
      }
      if (!d.objectStoreNames.contains('salidas')) {
        const s = d.createObjectStore('salidas', { keyPath: 'id', autoIncrement: true });
        s.createIndex('fecha_hora', 'fecha_hora');
      }
      if (!d.objectStoreNames.contains('papelera'))
        d.createObjectStore('papelera', { keyPath: 'id', autoIncrement: true });
    };
    req.onsuccess = (e) => { db = e.target.result; resolve(db); };
    req.onerror   = (e) => reject(e.target.error);
  });
}

function dbGetAll(store) {
  return new Promise((resolve, reject) => {
    const req = db.transaction(store, 'readonly').objectStore(store).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  });
}
function dbGet(store, id) {
  return new Promise((resolve, reject) => {
    const req = db.transaction(store, 'readonly').objectStore(store).get(id);
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  });
}
function dbAdd(store, data) {
  return new Promise((resolve, reject) => {
    const item = { ...data };
    if (item.id === undefined || item.id === null) delete item.id;
    const req = db.transaction(store, 'readwrite').objectStore(store).add(item);
    req.onsuccess = () => resolve({ ...item, id: req.result });
    req.onerror   = () => reject(req.error);
  });
}
function dbPut(store, data) {
  return new Promise((resolve, reject) => {
    const req = db.transaction(store, 'readwrite').objectStore(store).put(data);
    req.onsuccess = () => resolve(data);
    req.onerror   = () => reject(req.error);
  });
}
function dbDelete(store, id) {
  return new Promise((resolve, reject) => {
    const req = db.transaction(store, 'readwrite').objectStore(store).delete(id);
    req.onsuccess = () => resolve(true);
    req.onerror   = () => reject(req.error);
  });
}

// ──────────────────────────────────────────────
// ESTADO GLOBAL
// ──────────────────────────────────────────────
let inventario  = [];
let movimientos = [];
let papelera    = [];
let modoEdicion = false;

// ──────────────────────────────────────────────
// INIT
// ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  iniciarReloj();
  dibujarLogos();

  // FIX #2: tipo de salida reactivo
  document.getElementById('salida-tipo').addEventListener('change', onTipoSalidaChange);
  // Recalcular costo total al cambiar cantidad
  document.getElementById('entrada-cantidad').addEventListener('input', calcularCostoTotal);

  try {
    await abrirDB();
    await cargarTodo();
  } catch (err) {
    console.error('Error al abrir base de datos:', err);
    toast('Error al iniciar la base de datos', 'error');
  }

  document.getElementById('reporte-fecha').value = hoy();
  document.getElementById('reporte-fecha').addEventListener('change', cargarReporte);
  document.getElementById('reporte-tipo').addEventListener('change', cargarReporte);
  cargarReporte();
});

async function cargarTodo() {
  const [prods, entradas, salidas, pap] = await Promise.all([
    dbGetAll('productos'),
    dbGetAll('entradas'),
    dbGetAll('salidas'),
    dbGetAll('papelera')
  ]);

  inventario = prods;
  papelera   = pap;

  movimientos = [
    ...entradas.map(e => ({ ...e, tipo: 'entrada' })),
    ...salidas.map(s => ({ ...s, tipo: 'salida'  }))
  ].sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));

  actualizarDashboard();
  renderTablaInventario();
  renderEntradasHoy();
  renderSalidasHoy();
  llenarSelectProductos();
  renderPapelera();
}

// ──────────────────────────────────────────────
// LOGO ALTO BALANCE (canvas)
// ──────────────────────────────────────────────
function dibujarLogos() {
  ['logo-sidebar', 'logo-header'].forEach(id => {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const size = canvas.width;
    const ctx  = canvas.getContext('2d');
    const s    = size / 32;
    const cx   = size / 2;
    ctx.clearRect(0, 0, size, size);
    // "A" azul marino
    ctx.fillStyle = '#1e3a6e';
    ctx.beginPath();
    ctx.moveTo(cx, 3 * s);
    ctx.lineTo(cx - 11 * s, 26 * s);
    ctx.lineTo(cx - 6 * s, 26 * s);
    ctx.lineTo(cx, 12 * s);
    ctx.lineTo(cx + 6 * s, 26 * s);
    ctx.lineTo(cx + 11 * s, 26 * s);
    ctx.closePath();
    ctx.fill();
    // Hueco interior
    ctx.fillStyle = '#0f1208';
    ctx.beginPath();
    ctx.moveTo(cx, 16 * s);
    ctx.lineTo(cx - 4 * s, 23 * s);
    ctx.lineTo(cx + 4 * s, 23 * s);
    ctx.closePath();
    ctx.fill();
    // Flecha punta
    ctx.fillStyle = '#1e3a6e';
    ctx.beginPath();
    ctx.moveTo(cx, 0);
    ctx.lineTo(cx - 3 * s, 5 * s);
    ctx.lineTo(cx + 3 * s, 5 * s);
    ctx.closePath();
    ctx.fill();
    // Curva dorada
    ctx.strokeStyle = '#b8962e';
    ctx.lineWidth   = 2.5 * s;
    ctx.lineCap     = 'round';
    ctx.beginPath();
    ctx.moveTo(cx - 12 * s, 20 * s);
    ctx.bezierCurveTo(cx - 4 * s, 28 * s, cx + 8 * s, 28 * s, cx + 14 * s, 18 * s);
    ctx.stroke();
  });
}

// ──────────────────────────────────────────────
// RELOJ
// ──────────────────────────────────────────────
function iniciarReloj() {
  function tick() {
    const now = new Date();
    document.getElementById('clock').textContent = now.toLocaleTimeString('es-AR');
    document.getElementById('dateDisplay').textContent =
      now.toLocaleDateString('es-AR', { day:'2-digit', month:'2-digit', year:'numeric' });
  }
  tick();
  setInterval(tick, 1000);
}

// ──────────────────────────────────────────────
// NAVEGACIÓN
// ──────────────────────────────────────────────
function navigate(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById(`page-${page}`).classList.add('active');
  document.querySelector(`[data-page="${page}"]`).classList.add('active');
  document.getElementById('sidebar').classList.remove('open');
  if (page === 'entradas') renderEntradasHoy();
  if (page === 'salidas')  renderSalidasHoy();
  if (page === 'reportes') cargarReporte();
  if (page === 'papelera') renderPapelera();
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

// ──────────────────────────────────────────────
// ESTADO DE VENCIMIENTO
// ──────────────────────────────────────────────
function calcularEstado(fechaVenc) {
  if (!fechaVenc) return 'novence';
  const hoyDate = new Date(); hoyDate.setHours(0,0,0,0);
  const venc    = new Date(fechaVenc + 'T00:00:00');
  const diff    = Math.floor((venc - hoyDate) / 86400000);
  if (diff < 0)   return 'vencido';
  if (diff <= 30) return 'por_vencer';
  return 'ok';
}

function diasParaVencer(fechaVenc) {
  if (!fechaVenc) return null;
  const hoyDate = new Date(); hoyDate.setHours(0,0,0,0);
  return Math.floor((new Date(fechaVenc + 'T00:00:00') - hoyDate) / 86400000);
}

function estadoBadge(estado) {
  const mapa = {
    ok:        '<span class="status status-ok">Vigente</span>',
    por_vencer:'<span class="status status-warn">Por Vencer</span>',
    vencido:   '<span class="status status-danger">Vencido</span>',
    novence:   '<span class="status status-novence">Sin vencimiento</span>'
  };
  return mapa[estado] || mapa.ok;
}

// ──────────────────────────────────────────────
// DASHBOARD
// ──────────────────────────────────────────────
function actualizarDashboard() {
  const vencidos  = inventario.filter(p => calcularEstado(p.fecha_vencimiento) === 'vencido');
  const porVencer = inventario.filter(p => calcularEstado(p.fecha_vencimiento) === 'por_vencer');
  const ok        = inventario.filter(p => ['ok','novence'].includes(calcularEstado(p.fecha_vencimiento)));

  document.getElementById('stat-total').textContent   = inventario.length;
  document.getElementById('stat-ok').textContent      = ok.length;
  document.getElementById('stat-warn').textContent    = porVencer.length;
  document.getElementById('stat-danger').textContent  = vencidos.length;
  document.getElementById('badge-warn-count').textContent = porVencer.length;
  document.getElementById('badge-venc-count').textContent = vencidos.length;

  // Valor total del inventario disponible (no vencidos, stock > 0)
  const valorTotal = inventario
    .filter(p => calcularEstado(p.fecha_vencimiento) !== 'vencido' && parseFloat(p.cantidad) > 0)
    .reduce((s, p) => s + (parseFloat(p.costo_unitario)||0) * (parseFloat(p.cantidad)||0), 0);
  document.getElementById('stat-valor').textContent = `Bs ${formatNum(valorTotal)}`;

  // FIX #1: Solo vigentes (ok + novence + por_vencer) con cantidad > 0
  renderTablaDisponibles();

  renderMiniTabla('tbl-por-vencer',
    [...porVencer].sort((a,b) => new Date(a.fecha_vencimiento) - new Date(b.fecha_vencimiento)).slice(0,8),
    row => `
      <td>${escHtml(row.nombre)}</td>
      <td class="mono">${escHtml(row.lote||'—')}</td>
      <td>${row.cantidad} ${row.unidad||''}</td>
      <td class="mono">${formatFecha(row.fecha_vencimiento)}</td>
      <td><span class="badge badge-warn">${diasParaVencer(row.fecha_vencimiento)}d</span></td>
    `
  );

  renderMiniTabla('tbl-vencidos', vencidos.slice(0,8), row => `
    <td>${escHtml(row.nombre)}</td>
    <td class="mono">${escHtml(row.lote||'—')}</td>
    <td>${row.cantidad} ${row.unidad||''}</td>
    <td class="mono">${formatFecha(row.fecha_vencimiento)}</td>
  `);

  const recientes = [...movimientos].slice(0,10);
  renderMiniTabla('tbl-movimientos', recientes, row => `
    <td class="mono">${formatHora(row.fecha_hora)}</td>
    <td><span class="type-${row.subtipo === 'merma' ? 'merma' : row.tipo}">${
      row.tipo === 'entrada' ? '↓ Entrada'
      : row.subtipo === 'merma' ? '⚠ Merma'
      : '↑ Salida'
    }</span></td>
    <td>${escHtml(row.producto_nombre||'—')}</td>
    <td>${row.cantidad} ${row.unidad||''}</td>
    <td class="mono">${escHtml(row.lote||'—')}</td>
    <td>
      <button class="btn btn-danger" style="font-size:11px;padding:3px 8px"
        onclick="eliminarMovimiento('${row.tipo}', ${row.id})">✕ Anular</button>
    </td>
  `);
}

// FIX #1: Tabla disponibles — excluye vencidos y cantidad = 0
function renderTablaDisponibles() {
  const tbody = document.querySelector('#tbl-disponibles tbody');

  // Solo productos NO vencidos y con stock > 0
  const disponibles = inventario.filter(p => {
    const estado = calcularEstado(p.fecha_vencimiento);
    return estado !== 'vencido' && parseFloat(p.cantidad) > 0;
  });

  document.getElementById('badge-disponibles').textContent = disponibles.length;

  if (!disponibles.length) {
    tbody.innerHTML = `<tr><td colspan="9" class="table-empty">Sin productos disponibles en este momento</td></tr>`;
    return;
  }

  tbody.innerHTML = disponibles.map(p => {
    const comprado = movimientos
      .filter(m => m.tipo === 'entrada' && m.producto_id === p.id)
      .reduce((s, m) => s + (parseFloat(m.cantidad)||0), 0);
    const vendido = movimientos
      .filter(m => m.tipo === 'salida' && m.producto_id === p.id)
      .reduce((s, m) => s + (parseFloat(m.cantidad)||0), 0);
    const estado = calcularEstado(p.fecha_vencimiento);
    const costoUnit  = parseFloat(p.costo_unitario) || 0;
    const valorProd  = costoUnit * parseFloat(p.cantidad);
    return `
      <tr>
        <td><strong>${escHtml(p.nombre)}</strong><br><small style="color:var(--text3)">${escHtml(p.categoria||'')}</small></td>
        <td class="mono">${escHtml(p.lote||'—')}</td>
        <td class="mono" style="color:var(--green-l)">${comprado||'—'}</td>
        <td class="mono" style="color:var(--red-l)">${vendido||'—'}</td>
        <td class="mono" style="font-weight:700;font-size:15px">${p.cantidad}</td>
        <td class="mono">${escHtml(p.unidad||'—')}</td>
        <td class="mono" style="color:#c09af0">${costoUnit > 0 ? formatBs(costoUnit) : '—'}</td>
        <td class="mono" style="color:#c09af0;font-weight:600">${costoUnit > 0 ? formatBs(valorProd) : '—'}</td>
        <td class="mono">${formatFecha(p.fecha_fabricacion)}</td>
        <td class="mono">${formatFecha(p.fecha_vencimiento)}</td>
        <td>${estadoBadge(estado)}</td>
      </tr>
    `;
  }).join('');
}

function renderMiniTabla(id, data, rowFn) {
  const tbody = document.querySelector(`#${id} tbody`);
  if (!tbody) return;
  tbody.innerHTML = data.length
    ? data.map(row => `<tr>${rowFn(row)}</tr>`).join('')
    : `<tr><td colspan="10" class="table-empty">Sin registros</td></tr>`;
}

// ──────────────────────────────────────────────
// INVENTARIO
// ──────────────────────────────────────────────
function renderTablaInventario(filtro = '') {
  const tbody = document.querySelector('#tbl-inventario tbody');
  const datos = inventario.filter(p =>
    !filtro || p.nombre.toLowerCase().includes(filtro.toLowerCase())
  );
  if (!datos.length) {
    tbody.innerHTML = `<tr><td colspan="8" class="table-empty">Sin productos${filtro?' con ese filtro':''}</td></tr>`;
    return;
  }
  tbody.innerHTML = datos.map(p => {
    const estado     = calcularEstado(p.fecha_vencimiento);
    const costoUnit  = parseFloat(p.costo_unitario) || 0;
    const valorProd  = costoUnit * parseFloat(p.cantidad);
    return `
      <tr>
        <td><strong>${escHtml(p.nombre)}</strong><br><small style="color:var(--text3)">${escHtml(p.categoria||'')}</small></td>
        <td class="mono" style="font-size:15px;font-weight:600">${p.cantidad}</td>
        <td class="mono">${escHtml(p.unidad||'—')}</td>
        <td class="mono">${escHtml(p.lote||'—')}</td>
        <td class="mono" style="color:#c09af0">${costoUnit > 0 ? formatBs(costoUnit) : '—'}</td>
        <td class="mono" style="color:#c09af0;font-weight:600">${costoUnit > 0 ? formatBs(valorProd) : '—'}</td>
        <td class="mono">${formatFecha(p.fecha_fabricacion)}</td>
        <td class="mono">${formatFecha(p.fecha_vencimiento)}</td>
        <td>${estadoBadge(estado)}</td>
        <td>
          <div class="action-btns">
            <button class="btn btn-edit" onclick="editarProducto(${p.id})">✎ Editar</button>
            <button class="btn btn-danger" onclick="eliminarProducto(${p.id})">✕</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function filtrarInventario() {
  renderTablaInventario(document.getElementById('search-inventario').value);
}

// ──────────────────────────────────────────────
// Calcular costo total en tiempo real al escribir costo unitario o cantidad
function calcularCostoTotal() {
  const cant  = parseFloat(document.getElementById('entrada-cantidad').value) || 0;
  const cunit = parseFloat(document.getElementById('entrada-costo-unitario').value) || 0;
  const total = cant * cunit;
  document.getElementById('entrada-costo-total').value =
    total > 0 ? `Bs ${formatNum(total)}` : '';
}

// ──────────────────────────────────────────────
// FIX #3: VALIDACIÓN DE FECHAS (fabricación / vencimiento)
// Límite legal: 2 años desde fabricación
// ──────────────────────────────────────────────
function calcularVencimientoEntrada() {
  const fabVal  = document.getElementById('entrada-fabricacion').value;
  const vencEl  = document.getElementById('entrada-vencimiento');
  const hintEl  = document.getElementById('hint-vencimiento-entrada');
  if (!fabVal) return;

  const fab     = new Date(fabVal + 'T00:00:00');
  const maxVenc = new Date(fab);
  maxVenc.setFullYear(maxVenc.getFullYear() + 2);
  const maxStr  = maxVenc.toISOString().slice(0,10);

  vencEl.max   = maxStr;
  vencEl.value = maxStr; // sugiere la fecha máxima
  vencEl.classList.remove('input-error', 'input-warn');
  hintEl.textContent = `Vencimiento máximo según ley: ${formatFecha(maxStr)}`;
  hintEl.className   = 'field-hint hint-warn';
}

function validarVencimientoEntrada() {
  const fabVal  = document.getElementById('entrada-fabricacion').value;
  const vencVal = document.getElementById('entrada-vencimiento').value;
  const vencEl  = document.getElementById('entrada-vencimiento');
  const hintEl  = document.getElementById('hint-vencimiento-entrada');
  if (!fabVal || !vencVal) return;

  const fab     = new Date(fabVal  + 'T00:00:00');
  const venc    = new Date(vencVal + 'T00:00:00');
  const maxVenc = new Date(fab);
  maxVenc.setFullYear(maxVenc.getFullYear() + 2);

  if (venc > maxVenc) {
    vencEl.classList.add('input-error');
    hintEl.textContent = `⚠ Excede el límite legal de 2 años (máx: ${formatFecha(maxVenc.toISOString().slice(0,10))})`;
    hintEl.className   = 'field-hint hint-error';
  } else {
    vencEl.classList.remove('input-error');
    vencEl.classList.add('input-warn');
    const diasRestantes = Math.floor((maxVenc - venc) / 86400000);
    hintEl.textContent = `✓ Dentro del límite legal · ${diasRestantes} días antes del máximo`;
    hintEl.className   = 'field-hint hint-warn';
  }
}

function calcularVencimientoModal() {
  const fabVal  = document.getElementById('prod-fabricacion').value;
  const vencEl  = document.getElementById('prod-vencimiento');
  const hintEl  = document.getElementById('hint-vencimiento-modal');
  if (!fabVal) return;

  const fab     = new Date(fabVal + 'T00:00:00');
  const maxVenc = new Date(fab);
  maxVenc.setFullYear(maxVenc.getFullYear() + 2);
  const maxStr  = maxVenc.toISOString().slice(0,10);

  vencEl.max   = maxStr;
  vencEl.value = maxStr;
  vencEl.classList.remove('input-error');
  hintEl.textContent = `Vencimiento máximo según ley: ${formatFecha(maxStr)}`;
  hintEl.className   = 'field-hint hint-warn';
}

function validarVencimientoModal() {
  const fabVal  = document.getElementById('prod-fabricacion').value;
  const vencVal = document.getElementById('prod-vencimiento').value;
  const vencEl  = document.getElementById('prod-vencimiento');
  const hintEl  = document.getElementById('hint-vencimiento-modal');
  if (!fabVal || !vencVal) return;

  const fab     = new Date(fabVal  + 'T00:00:00');
  const venc    = new Date(vencVal + 'T00:00:00');
  const maxVenc = new Date(fab);
  maxVenc.setFullYear(maxVenc.getFullYear() + 2);

  if (venc > maxVenc) {
    vencEl.classList.add('input-error');
    hintEl.textContent = `⚠ Excede el límite legal de 2 años (máx: ${formatFecha(maxVenc.toISOString().slice(0,10))})`;
    hintEl.className   = 'field-hint hint-error';
  } else {
    vencEl.classList.remove('input-error');
    hintEl.textContent = `✓ Dentro del límite legal`;
    hintEl.className   = 'field-hint hint-warn';
  }
}

// Mostrar/ocultar campos de fechas según categoría
function onCategoriaChange() {
  // Herramientas y Otros no tienen vencimiento
  const cat      = document.getElementById('prod-categoria').value;
  const sinVenc  = ['Herramientas', 'Otros'];
  const grupo    = document.getElementById('prod-fechas-group');
  if (sinVenc.includes(cat)) {
    grupo.style.opacity = '0.4';
    grupo.style.pointerEvents = 'none';
    document.getElementById('prod-fabricacion').value = '';
    document.getElementById('prod-vencimiento').value = '';
  } else {
    grupo.style.opacity = '1';
    grupo.style.pointerEvents = 'auto';
  }
}

// ──────────────────────────────────────────────
// MODAL PRODUCTO
// ──────────────────────────────────────────────
function abrirModalProducto(prod = null) {
  modoEdicion = !!prod;
  document.getElementById('modal-titulo').textContent   = modoEdicion ? 'Editar Producto' : 'Nuevo Producto';
  document.getElementById('prod-id').value              = prod ? prod.id : '';
  document.getElementById('prod-nombre').value          = prod ? prod.nombre : '';
  document.getElementById('prod-cantidad').value        = prod ? prod.cantidad : '';
  document.getElementById('prod-lote').value            = prod ? (prod.lote||'') : '';
  document.getElementById('prod-fabricacion').value     = prod ? (prod.fecha_fabricacion||'') : '';
  document.getElementById('prod-vencimiento').value     = prod ? (prod.fecha_vencimiento||'') : '';
  document.getElementById('prod-unidad').value          = prod ? (prod.unidad||'kg') : 'kg';
  document.getElementById('prod-categoria').value       = prod ? (prod.categoria||'Agroquímicos') : 'Agroquímicos';
  document.getElementById('modal-producto').classList.add('open');
  onCategoriaChange();
  setTimeout(() => document.getElementById('prod-nombre').focus(), 100);
}

function cerrarModal(id) {
  document.getElementById(id).classList.remove('open');
}

async function guardarProducto() {
  const nombre      = document.getElementById('prod-nombre').value.trim();
  const cantidad    = parseFloat(document.getElementById('prod-cantidad').value);
  const lote        = document.getElementById('prod-lote').value.trim();
  const fabricacion = document.getElementById('prod-fabricacion').value || null;
  const venc        = document.getElementById('prod-vencimiento').value || null;
  const unidad      = document.getElementById('prod-unidad').value;
  const categoria   = document.getElementById('prod-categoria').value;
  const idStr       = document.getElementById('prod-id').value;

  if (!nombre || isNaN(cantidad) || !lote) {
    toast('Completa nombre, cantidad y lote', 'error'); return;
  }

  // Validar límite de 2 años si hay fabricación y vencimiento
  if (fabricacion && venc) {
    const fab     = new Date(fabricacion + 'T00:00:00');
    const vencD   = new Date(venc + 'T00:00:00');
    const maxVenc = new Date(fab);
    maxVenc.setFullYear(maxVenc.getFullYear() + 2);
    if (vencD > maxVenc) {
      toast('La fecha de vencimiento supera los 2 años desde fabricación (límite legal)', 'error');
      return;
    }
  }

  const prodData = { nombre, cantidad, lote, fecha_fabricacion: fabricacion, fecha_vencimiento: venc, unidad, categoria };

  if (modoEdicion && idStr) {
    const id = parseInt(idStr);
    await dbPut('productos', { ...prodData, id });
    const idx = inventario.findIndex(p => p.id === id);
    if (idx >= 0) inventario[idx] = { ...prodData, id };
    toast('Producto actualizado ✓', 'ok');
  } else {
    const nuevo = await dbAdd('productos', prodData);
    inventario.push(nuevo);
    toast('Producto registrado ✓', 'ok');
  }

  cerrarModal('modal-producto');
  actualizarDashboard();
  renderTablaInventario();
  llenarSelectProductos();
}

function editarProducto(id) {
  const prod = inventario.find(p => p.id === id);
  if (prod) abrirModalProducto(prod);
}

async function eliminarProducto(id) {
  const prod = inventario.find(p => p.id === id);
  if (!prod) return;
  if (!confirm(`¿Eliminar "${prod.nombre}" (Lote: ${prod.lote||'—'})?\nEsta acción no se puede deshacer.`)) return;
  await dbDelete('productos', id);
  inventario = inventario.filter(p => p.id !== id);
  actualizarDashboard();
  renderTablaInventario();
  llenarSelectProductos();
  toast('Producto eliminado', 'warn');
}

// ──────────────────────────────────────────────
// SELECTS
// ──────────────────────────────────────────────
function llenarSelectProductos() {
  ['entrada-producto','salida-producto'].forEach(selId => {
    const sel = document.getElementById(selId);
    const val = sel.value;
    sel.innerHTML = '<option value="">— Seleccionar producto —</option>' +
      inventario.map(p =>
        `<option value="${p.id}">${escHtml(p.nombre)} · Lote ${escHtml(p.lote||'?')} [${p.cantidad} ${p.unidad||''}]</option>`
      ).join('');
    if (val) sel.value = val;
  });
}

function onProductoEntradaChange() {
  const id   = parseInt(document.getElementById('entrada-producto').value);
  const prod = inventario.find(p => p.id === id);
  if (prod && !document.getElementById('entrada-nombre').value)
    document.getElementById('entrada-nombre').value = prod.nombre;
}

function onProductoSalidaChange() {
  const id   = parseInt(document.getElementById('salida-producto').value);
  const prod = inventario.find(p => p.id === id);
  const info = document.getElementById('salida-stock-info');
  if (prod) {
    const costoUnit  = parseFloat(prod.costo_unitario) || 0;
    const costoTotal = costoUnit * parseFloat(prod.cantidad);
    info.classList.remove('hidden');
    info.innerHTML = `
      Stock: <strong>${prod.cantidad} ${prod.unidad||''}</strong> &nbsp;|&nbsp;
      Lote: <strong>${prod.lote||'—'}</strong> &nbsp;|&nbsp;
      Vence: <strong>${prod.fecha_vencimiento ? formatFecha(prod.fecha_vencimiento) : 'Sin vencimiento'}</strong> &nbsp;|&nbsp;
      ${estadoBadge(calcularEstado(prod.fecha_vencimiento))}
      ${costoUnit > 0 ? `&nbsp;|&nbsp; Costo unit.: <strong style="color:#c09af0">${formatBs(costoUnit)}</strong> &nbsp;|&nbsp; Valor en stock: <strong style="color:#c09af0">${formatBs(costoTotal)}</strong>` : ''}
    `;
  } else {
    info.classList.add('hidden');
  }
}

// FIX #2: Panel dinámico cliente/merma según tipo de salida
function onTipoSalidaChange() {
  const tipo         = document.getElementById('salida-tipo').value;
  const clienteGroup = document.getElementById('salida-cliente-group');
  const mermaGroup   = document.getElementById('salida-merma-group');
  if (tipo === 'merma') {
    clienteGroup.classList.add('hidden');
    mermaGroup.classList.remove('hidden');
  } else {
    clienteGroup.classList.remove('hidden');
    mermaGroup.classList.add('hidden');
  }
}

// ──────────────────────────────────────────────
// ENTRADAS (con fecha de fabricación — FIX #3)
// ──────────────────────────────────────────────
async function registrarEntrada() {
  const prodId      = parseInt(document.getElementById('entrada-producto').value) || null;
  const nombre      = document.getElementById('entrada-nombre').value.trim();
  const cantidad    = parseFloat(document.getElementById('entrada-cantidad').value);
  const lote        = document.getElementById('entrada-lote').value.trim();
  const fabricacion = document.getElementById('entrada-fabricacion').value || null;
  const venc        = document.getElementById('entrada-vencimiento').value || null;
  const unidad      = document.getElementById('entrada-unidad').value;
  const proveedor   = document.getElementById('entrada-proveedor').value.trim();
  const obs         = document.getElementById('entrada-obs').value.trim();
  const costoUnit   = parseFloat(document.getElementById('entrada-costo-unitario').value) || 0;

  if (!nombre && !prodId) { toast('Selecciona o escribe el nombre del producto', 'error'); return; }
  if (!cantidad || cantidad <= 0) { toast('Ingresa una cantidad válida', 'error'); return; }
  if (!lote) { toast('El número de lote es obligatorio', 'error'); return; }

  // Validar límite legal de 2 años
  if (fabricacion && venc) {
    const fab     = new Date(fabricacion + 'T00:00:00');
    const vencD   = new Date(venc + 'T00:00:00');
    const maxVenc = new Date(fab);
    maxVenc.setFullYear(maxVenc.getFullYear() + 2);
    if (vencD > maxVenc) {
      toast('La fecha de vencimiento supera los 2 años desde fabricación (límite legal)', 'error');
      return;
    }
  }

  let prod = prodId ? inventario.find(p => p.id === prodId) : null;
  const nombreFinal = nombre || prod?.nombre || '';

  if (prod) {
    if (prod.lote === lote && prod.fecha_vencimiento === venc) {
      prod.cantidad = (parseFloat(prod.cantidad)||0) + cantidad;
      if (fabricacion) prod.fecha_fabricacion = fabricacion;
      // Actualizar costo unitario si se ingresó uno nuevo
      if (costoUnit > 0) prod.costo_unitario = costoUnit;
      await dbPut('productos', prod);
      toast(`Stock sumado al lote existente ✓`, 'ok');
    } else {
      const nuevo = await dbAdd('productos', {
        nombre: nombreFinal, cantidad, lote,
        fecha_fabricacion: fabricacion, fecha_vencimiento: venc,
        unidad, categoria: prod.categoria||'Otros',
        costo_unitario: costoUnit
      });
      prod = nuevo;
      inventario.push(prod);
      toast(`Nuevo lote creado: ${nombreFinal} · ${lote} ✓`, 'ok');
    }
  } else {
    const nuevo = await dbAdd('productos', {
      nombre: nombreFinal, cantidad, lote,
      fecha_fabricacion: fabricacion, fecha_vencimiento: venc,
      unidad, categoria: 'Otros',
      costo_unitario: costoUnit
    });
    prod = nuevo;
    inventario.push(prod);
    toast(`Producto nuevo registrado ✓`, 'ok');
  }

  await dbAdd('entradas', {
    producto_id: prod.id, producto_nombre: prod.nombre,
    cantidad, unidad, lote,
    fecha_fabricacion: fabricacion, fecha_vencimiento: venc,
    costo_unitario: costoUnit,
    costo_total: costoUnit * cantidad,
    proveedor, observaciones: obs,
    fecha_hora: new Date().toISOString()
  });

  ['entrada-producto','entrada-nombre','entrada-cantidad','entrada-lote',
   'entrada-fabricacion','entrada-vencimiento','entrada-proveedor','entrada-obs',
   'entrada-costo-unitario','entrada-costo-total'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('hint-vencimiento-entrada').textContent = 'Dejar vacío si el producto no vence';
  document.getElementById('hint-vencimiento-entrada').className   = 'field-hint';

  await cargarTodo();
  cargarReporte();
}

function renderEntradasHoy() {
  const hoyStr      = hoy();
  const entradasHoy = movimientos
    .filter(m => m.tipo === 'entrada' && fechaLocalDeISO(m.fecha_hora) === hoyStr);
  document.getElementById('count-entradas-hoy').textContent = entradasHoy.length;
  const tbody = document.querySelector('#tbl-entradas-hoy tbody');
  if (!entradasHoy.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="table-empty">Sin entradas hoy</td></tr>`; return;
  }
  tbody.innerHTML = entradasHoy.map(row => `
    <tr>
      <td class="mono">${formatHora(row.fecha_hora)}</td>
      <td>${escHtml(row.producto_nombre||'—')}</td>
      <td>${row.cantidad} ${row.unidad||''}</td>
      <td class="mono">${escHtml(row.lote||'—')}</td>
      <td>
        <button class="btn btn-danger" style="font-size:11px;padding:3px 8px"
          onclick="eliminarMovimiento('entrada', ${row.id})">✕ Anular</button>
      </td>
    </tr>
  `).join('');
}

// ──────────────────────────────────────────────
// SALIDAS — FIX #2: mermas con motivo propio
// ──────────────────────────────────────────────
async function registrarSalida() {
  const prodId   = parseInt(document.getElementById('salida-producto').value);
  const cantidad = parseFloat(document.getElementById('salida-cantidad').value);
  const tipo     = document.getElementById('salida-tipo').value;
  const cliente  = document.getElementById('salida-cliente').value.trim();
  const mermaMotivo = document.getElementById('salida-merma-motivo').value;
  const obs      = document.getElementById('salida-obs').value.trim();

  if (!prodId)  { toast('Selecciona un producto', 'error'); return; }
  if (!cantidad || cantidad <= 0) { toast('Ingresa una cantidad válida', 'error'); return; }

  const prod = inventario.find(p => p.id === prodId);
  if (!prod) { toast('Producto no encontrado', 'error'); return; }

  // Bloquear venta (no merma) de productos vencidos
  if (calcularEstado(prod.fecha_vencimiento) === 'vencido' && tipo === 'venta') {
    toast(`⚠ No se puede vender "${prod.nombre}": producto VENCIDO. Usa "Merma / Pérdida" para registrar su baja.`, 'error');
    return;
  }

  if (cantidad > prod.cantidad) {
    toast(`Stock insuficiente. Disponible: ${prod.cantidad} ${prod.unidad||''}`, 'error');
    return;
  }

  prod.cantidad = parseFloat(prod.cantidad) - cantidad;
  await dbPut('productos', prod);

  // Referencia según tipo
  const referencia = tipo === 'merma' ? mermaMotivo : cliente;

  await dbAdd('salidas', {
    producto_id: prod.id, producto_nombre: prod.nombre,
    cantidad, unidad: prod.unidad, lote: prod.lote,
    subtipo: tipo, cliente: referencia, observaciones: obs,
    costo_unitario: parseFloat(prod.costo_unitario) || 0,
    costo_total: (parseFloat(prod.costo_unitario) || 0) * cantidad,
    fecha_hora: new Date().toISOString()
  });

  document.getElementById('salida-producto').value = '';
  document.getElementById('salida-cantidad').value = '';
  document.getElementById('salida-cliente').value  = '';
  document.getElementById('salida-obs').value      = '';
  document.getElementById('salida-stock-info').classList.add('hidden');
  document.getElementById('salida-tipo').value     = 'venta';
  onTipoSalidaChange(); // resetear panel

  const tipoLabel = tipo === 'merma' ? `Merma (${mermaMotivo})` : tipo;
  toast(`${tipoLabel} registrada: ${cantidad} ${prod.unidad||''} de "${prod.nombre}" ✓`, 'ok');
  await cargarTodo();
  cargarReporte();
}

function renderSalidasHoy() {
  const hoyStr     = hoy();
  const salidasHoy = movimientos
    .filter(m => m.tipo === 'salida' && fechaLocalDeISO(m.fecha_hora) === hoyStr);
  document.getElementById('count-salidas-hoy').textContent = salidasHoy.length;
  const tbody = document.querySelector('#tbl-salidas-hoy tbody');
  if (!salidasHoy.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="table-empty">Sin salidas hoy</td></tr>`; return;
  }
  tbody.innerHTML = salidasHoy.map(row => {
    const esMerma   = row.subtipo === 'merma';
    const tipoClass = esMerma ? 'type-merma' : 'type-salida';
    const tipoLabel = esMerma ? `⚠ Merma` : `↑ ${row.subtipo||'venta'}`;
    const ct = parseFloat(row.costo_total) || (parseFloat(row.costo_unitario)||0) * parseFloat(row.cantidad);
    return `
      <tr>
        <td class="mono">${formatHora(row.fecha_hora)}</td>
        <td>${escHtml(row.producto_nombre||'—')}</td>
        <td>${row.cantidad} ${row.unidad||''}</td>
        <td><span class="${tipoClass}">${tipoLabel}</span></td>
        <td class="mono" style="color:#c09af0;font-weight:600">${ct > 0 ? formatBs(ct) : '—'}</td>
        <td>
          <button class="btn btn-danger" style="font-size:11px;padding:3px 8px"
            onclick="eliminarMovimiento('salida', ${row.id})">✕ Anular</button>
        </td>
      </tr>
    `;
  }).join('');
}

// ──────────────────────────────────────────────
// ELIMINAR MOVIMIENTO → PAPELERA
// ──────────────────────────────────────────────
async function eliminarMovimiento(tipo, id) {
  const store = tipo === 'entrada' ? 'entradas' : 'salidas';
  const mov   = await dbGet(store, id);
  if (!mov) return;

  const efecto = tipo === 'entrada'
    ? `Se descontarán ${mov.cantidad} ${mov.unidad||''} del stock.`
    : `Se devolverán ${mov.cantidad} ${mov.unidad||''} al stock.`;

  if (!confirm(`¿Anular este movimiento?\n\nProducto: ${mov.producto_nombre}\nCantidad: ${mov.cantidad} ${mov.unidad||''}\nHora: ${formatHora(mov.fecha_hora)}\n\n⚠ ${efecto}\n\nEl registro irá a la Papelera.`)) return;

  const prod = inventario.find(p => p.id === mov.producto_id);
  if (prod) {
    prod.cantidad = tipo === 'entrada'
      ? Math.max(0, parseFloat(prod.cantidad) - mov.cantidad)
      : parseFloat(prod.cantidad) + mov.cantidad;
    await dbPut('productos', prod);
  }

  await dbAdd('papelera', {
    ...mov, tipo_original: tipo, store_original: store,
    id_original: id, eliminado_en: new Date().toISOString()
  });

  await dbDelete(store, id);
  toast('Movimiento anulado → guardado en Papelera ✓', 'warn');
  await cargarTodo();
  cargarReporte();
}

// ──────────────────────────────────────────────
// PAPELERA
// ──────────────────────────────────────────────
async function recuperarDePapelera(papeleraId) {
  const item = papelera.find(p => p.id === papeleraId);
  if (!item) return;
  if (!confirm(`¿Recuperar este movimiento?\n\nProducto: ${item.producto_nombre}\nCantidad: ${item.cantidad} ${item.unidad||''}\n\n⚠ El stock se restaurará automáticamente.`)) return;

  const { id, id_original, tipo_original, store_original, eliminado_en, ...movData } = item;
  await dbAdd(store_original, movData);

  const prod = inventario.find(p => p.id === item.producto_id);
  if (prod) {
    prod.cantidad = tipo_original === 'entrada'
      ? parseFloat(prod.cantidad) + item.cantidad
      : Math.max(0, parseFloat(prod.cantidad) - item.cantidad);
    await dbPut('productos', prod);
  }

  await dbDelete('papelera', papeleraId);
  toast('Movimiento recuperado y stock restaurado ✓', 'ok');
  await cargarTodo();
  cargarReporte();
}

async function eliminarDePapelera(id) {
  if (!confirm('¿Eliminar definitivamente este registro? No se podrá recuperar.')) return;
  await dbDelete('papelera', id);
  papelera = papelera.filter(p => p.id !== id);
  renderPapelera();
  document.getElementById('papelera-count').textContent = papelera.length;
  toast('Registro eliminado definitivamente', 'warn');
}

async function vaciarPapelera() {
  if (!papelera.length) { toast('La papelera ya está vacía', 'warn'); return; }
  if (!confirm(`¿Vaciar la papelera? Se eliminarán ${papelera.length} registros permanentemente.`)) return;
  for (const item of papelera) await dbDelete('papelera', item.id);
  papelera = [];
  renderPapelera();
  document.getElementById('papelera-count').textContent = 0;
  toast('Papelera vaciada', 'warn');
}

function renderPapelera() {
  const tbody = document.querySelector('#tbl-papelera tbody');
  if (!tbody) return;
  document.getElementById('papelera-count').textContent = papelera.length;
  if (!papelera.length) {
    tbody.innerHTML = `<tr><td colspan="8" class="table-empty">La papelera está vacía</td></tr>`; return;
  }
  const sorted = [...papelera].sort((a,b) => new Date(b.eliminado_en) - new Date(a.eliminado_en));
  tbody.innerHTML = sorted.map(item => `
    <tr>
      <td class="mono">${formatFecha2(item.fecha_hora)}</td>
      <td class="mono">${formatHora(item.fecha_hora)}</td>
      <td><span class="type-${item.tipo_original}">${item.tipo_original==='entrada'?'↓ Entrada':'↑ Salida'}</span></td>
      <td>${escHtml(item.producto_nombre||'—')}</td>
      <td>${item.cantidad} ${item.unidad||''}</td>
      <td class="mono">${escHtml(item.lote||'—')}</td>
      <td class="mono" style="color:var(--text3)">${formatHora(item.eliminado_en)}</td>
      <td>
        <div class="action-btns">
          <button class="btn btn-recover" onclick="recuperarDePapelera(${item.id})">↩ Recuperar</button>
          <button class="btn btn-danger"  onclick="eliminarDePapelera(${item.id})">✕</button>
        </div>
      </td>
    </tr>
  `).join('');
}

// ──────────────────────────────────────────────
// REPORTES
// ──────────────────────────────────────────────
let reporteActual = [];

function cargarReporte() {
  const fechaInput = document.getElementById('reporte-fecha').value;
  const tipoFiltro = document.getElementById('reporte-tipo').value;

  reporteActual = movimientos.filter(m => {
    const coincideFecha = !fechaInput || fechaLocalDeISO(m.fecha_hora) === fechaInput;
    const coincideTipo  = !tipoFiltro || m.tipo === tipoFiltro;
    return coincideFecha && coincideTipo;
  });

  const tituloFecha = fechaInput
    ? new Date(fechaInput+'T12:00:00').toLocaleDateString('es-AR',{day:'2-digit',month:'long',year:'numeric'})
    : 'Todos los períodos';
  document.getElementById('reporte-titulo').textContent =
    `Movimientos — ${tituloFecha}${tipoFiltro?' · Solo '+tipoFiltro+'s':''}`;
  document.getElementById('reporte-count').textContent = `${reporteActual.length} registros`;

  const tbody = document.querySelector('#tbl-reporte tbody');
  if (!reporteActual.length) {
    tbody.innerHTML = `<tr><td colspan="9" class="table-empty">Sin movimientos para los filtros seleccionados</td></tr>`;
    return;
  }
  tbody.innerHTML = reporteActual.map(m => {
    const esMerma   = m.subtipo === 'merma';
    const tipoClass = esMerma ? 'type-merma' : `type-${m.tipo}`;
    const tipoLabel = m.tipo === 'entrada' ? '↓ Entrada' : esMerma ? '⚠ Merma' : '↑ Salida';
    const cUnit     = parseFloat(m.costo_unitario) || 0;
    const cTotal    = parseFloat(m.costo_total)    || (cUnit * parseFloat(m.cantidad));
    return `
      <tr>
        <td class="mono">${formatFecha2(m.fecha_hora)}</td>
        <td class="mono">${formatHora(m.fecha_hora)}</td>
        <td><span class="${tipoClass}">${tipoLabel}</span></td>
        <td>${escHtml(m.producto_nombre||'—')}</td>
        <td style="font-weight:600">${m.cantidad}</td>
        <td class="mono">${escHtml(m.unidad||'—')}</td>
        <td class="mono">${escHtml(m.lote||'—')}</td>
        <td class="mono" style="color:#c09af0">${cUnit  > 0 ? formatBs(cUnit)  : '—'}</td>
        <td class="mono" style="color:#c09af0;font-weight:600">${cTotal > 0 ? formatBs(cTotal) : '—'}</td>
        <td class="mono" style="color:var(--text3)">${escHtml(m.cliente||m.proveedor||m.subtipo||'—')}</td>
        <td>
          <button class="btn btn-danger" style="font-size:11px;padding:3px 8px"
            onclick="eliminarMovimiento('${m.tipo}', ${m.id})">✕ Anular</button>
        </td>
      </tr>
    `;
  }).join('');
}

// ──────────────────────────────────────────────
// FIX #4: EXPORTAR EXCEL — formato profesional con estilos
// ──────────────────────────────────────────────
function exportarExcel() {
  if (!reporteActual.length) { toast('No hay datos para exportar', 'warn'); return; }
  if (typeof XLSX === 'undefined') { exportarCSV(); return; }

  const fechaReporte  = document.getElementById('reporte-fecha').value || hoy();
  const totalEntradas = reporteActual.filter(m => m.tipo === 'entrada').length;
  const totalSalidas  = reporteActual.filter(m => m.tipo === 'salida').length;
  const totalMermas   = reporteActual.filter(m => m.subtipo === 'merma').length;
  const generado      = new Date().toLocaleString('es-AR');

  const costoTotalEntradas = reporteActual
    .filter(m => m.tipo === 'entrada')
    .reduce((s, m) => s + (parseFloat(m.costo_total) || (parseFloat(m.costo_unitario)||0) * parseFloat(m.cantidad)), 0);
  const valorInventario = inventario
    .filter(p => calcularEstado(p.fecha_vencimiento) !== 'vencido' && parseFloat(p.cantidad) > 0)
    .reduce((s, p) => s + (parseFloat(p.costo_unitario)||0) * parseFloat(p.cantidad), 0);

  // ── HOJA 1: Portada / Resumen ─────────────────
  const aoa_resumen = [
    ['AGROPECUARIA EL CHURQUI', '', '', ''],
    ['Sistema de Gestión de Inventario', '', '', ''],
    ['Gestionado por: Alto Balance Consultora', '', '', ''],
    ['', '', '', ''],
    ['INFORMACIÓN DEL REPORTE', '', '', ''],
    ['Fecha del reporte:', fechaReporte, '', ''],
    ['Período:', fechaReporte || 'Todos los períodos', '', ''],
    ['Generado:', generado, '', ''],
    ['', '', '', ''],
    ['RESUMEN DE MOVIMIENTOS', '', '', ''],
    ['Concepto', 'Cantidad', 'Monto (Bs)', ''],
    ['Total movimientos', reporteActual.length, '', ''],
    ['Entradas', totalEntradas, Number(costoTotalEntradas.toFixed(2)), ''],
    ['Salidas (ventas)', totalSalidas - totalMermas, '', ''],
    ['Mermas / Pérdidas', totalMermas, '', ''],
    ['', '', '', ''],
    ['RESUMEN FINANCIERO', '', '', ''],
    ['Valor total inventario disponible (Bs)', Number(valorInventario.toFixed(2)), '', ''],
    ['Productos con costo registrado', inventario.filter(p => parseFloat(p.costo_unitario) > 0).length, '', ''],
    ['', '', '', ''],
    ['INVENTARIO', '', '', ''],
    ['Total productos en stock', inventario.filter(p => parseFloat(p.cantidad) > 0).length, '', ''],
    ['Productos vencidos', inventario.filter(p => calcularEstado(p.fecha_vencimiento) === 'vencido').length, '', ''],
    ['Productos por vencer', inventario.filter(p => calcularEstado(p.fecha_vencimiento) === 'por_vencer').length, '', ''],
  ];
  const wsResumen = XLSX.utils.aoa_to_sheet(aoa_resumen);
  wsResumen['!cols'] = [{wch:38},{wch:22},{wch:18},{wch:10}];
  wsResumen['!merges'] = [
    { s:{r:0,c:0}, e:{r:0,c:3} },
    { s:{r:1,c:0}, e:{r:1,c:3} },
    { s:{r:2,c:0}, e:{r:2,c:3} },
  ];

  // ── HOJA 2: Movimientos ───────────────────────
  const encabezados = [
    'Fecha', 'Hora', 'Tipo', 'Subtipo', 'Producto',
    'Cantidad', 'Unidad', 'Lote', 'Costo Unit. (Bs)', 'Costo Total (Bs)', 'Referencia / Cliente', 'Observaciones'
  ];
  const filasMovimientos = reporteActual.map(m => {
    const cu = parseFloat(m.costo_unitario) || 0;
    const ct = parseFloat(m.costo_total)    || cu * parseFloat(m.cantidad);
    return [
      formatFecha2(m.fecha_hora),
      formatHora(m.fecha_hora),
      m.tipo === 'entrada' ? 'ENTRADA' : 'SALIDA',
      m.subtipo || '',
      m.producto_nombre || '',
      Number(m.cantidad),
      m.unidad || '',
      m.lote || '',
      cu > 0 ? Number(cu.toFixed(2)) : '',
      ct > 0 ? Number(ct.toFixed(2)) : '',
      m.cliente || m.proveedor || '',
      m.observaciones || ''
    ];
  });

  const wsMovimientos = XLSX.utils.aoa_to_sheet([encabezados, ...filasMovimientos]);
  wsMovimientos['!cols'] = [
    {wch:12},{wch:10},{wch:10},{wch:14},{wch:28},
    {wch:10},{wch:10},{wch:14},{wch:16},{wch:16},{wch:22},{wch:26}
  ];

  // ── HOJA 3: Inventario actual ─────────────────
  const encInv = ['Producto','Categoría','Cantidad','Unidad','Lote','Costo Unit. (Bs)','Valor Total (Bs)','Fabricación','Vencimiento','Estado'];
  const filasInv = inventario.map(p => {
    const cu = parseFloat(p.costo_unitario) || 0;
    const ct = cu * parseFloat(p.cantidad);
    return [
      p.nombre,
      p.categoria || '',
      Number(p.cantidad),
      p.unidad || '',
      p.lote || '',
      cu > 0 ? Number(cu.toFixed(2)) : '',
      ct > 0 ? Number(ct.toFixed(2)) : '',
      p.fecha_fabricacion ? formatFecha(p.fecha_fabricacion) : 'Sin dato',
      p.fecha_vencimiento ? formatFecha(p.fecha_vencimiento) : 'Sin vencimiento',
      calcularEstado(p.fecha_vencimiento) === 'ok'        ? 'Vigente'
      : calcularEstado(p.fecha_vencimiento) === 'por_vencer' ? 'Por Vencer'
      : calcularEstado(p.fecha_vencimiento) === 'vencido'    ? 'Vencido'
      : 'Sin vencimiento'
    ];
  });
  const wsInventario = XLSX.utils.aoa_to_sheet([encInv, ...filasInv]);
  wsInventario['!cols'] = [
    {wch:30},{wch:16},{wch:10},{wch:10},{wch:14},{wch:16},{wch:16},{wch:14},{wch:16},{wch:14}
  ];

  // ── Crear y descargar workbook ─────────────────
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsResumen,     'Resumen');
  XLSX.utils.book_append_sheet(wb, wsMovimientos, 'Movimientos');
  XLSX.utils.book_append_sheet(wb, wsInventario,  'Inventario Actual');

  XLSX.writeFile(wb, `churqui_reporte_${fechaReporte}.xlsx`);
  toast('Excel con 3 hojas descargado ✓', 'ok');
}

function exportarCSV() {
  const headers = ['Fecha','Hora','Tipo','Subtipo','Producto','Cantidad','Unidad','Lote','Referencia','Observaciones'];
  let csv = '\uFEFF' + headers.join(';') + '\r\n';
  reporteActual.forEach(m => {
    const f = [
      formatFecha2(m.fecha_hora), formatHora(m.fecha_hora),
      m.tipo, m.subtipo||'', m.producto_nombre||'',
      m.cantidad, m.unidad||'', m.lote||'',
      m.cliente||m.proveedor||'', m.observaciones||''
    ];
    csv += f.map(v => `"${String(v).replace(/"/g,'""')}"`).join(';') + '\r\n';
  });
  descargarArchivo(csv, `churqui_reporte_${hoy()}.csv`, 'text/csv;charset=utf-8');
  toast('Reporte CSV descargado ✓', 'ok');
}

// ──────────────────────────────────────────────
// EXPORTAR PDF
// ──────────────────────────────────────────────
function exportarPDF() {
  if (!reporteActual.length) { toast('No hay datos para exportar', 'warn'); return; }

  const fecha         = document.getElementById('reporte-fecha').value || hoy();
  const titulo        = `Reporte — ${new Date(fecha+'T12:00:00').toLocaleDateString('es-AR',{day:'2-digit',month:'long',year:'numeric'})}`;
  const totalEntradas = reporteActual.filter(m => m.tipo === 'entrada').length;
  const totalSalidas  = reporteActual.filter(m => m.tipo === 'salida').length;
  const totalMermas   = reporteActual.filter(m => m.subtipo === 'merma').length;

  const canvas  = document.getElementById('logo-header');
  const logoB64 = canvas ? canvas.toDataURL('image/png') : '';

  const filas = reporteActual.map(m => {
    const esMerma   = m.subtipo === 'merma';
    const color     = m.tipo === 'entrada' ? '#2d6a0a' : esMerma ? '#7a5200' : '#8b1a1a';
    const tipoLabel = m.tipo === 'entrada' ? '↓ ENTRADA' : esMerma ? '⚠ MERMA' : '↑ SALIDA';
    const cu = parseFloat(m.costo_unitario) || 0;
    const ct = parseFloat(m.costo_total)    || cu * parseFloat(m.cantidad);
    return `
      <tr>
        <td>${formatFecha2(m.fecha_hora)}</td>
        <td>${formatHora(m.fecha_hora)}</td>
        <td style="color:${color};font-weight:700">${tipoLabel}</td>
        <td>${escHtml(m.producto_nombre||'')}</td>
        <td><strong>${m.cantidad}</strong> ${m.unidad||''}</td>
        <td>${escHtml(m.lote||'—')}</td>
        <td style="color:#6a3db8">${cu > 0 ? 'Bs '+formatNum(cu) : '—'}</td>
        <td style="color:#6a3db8;font-weight:600">${ct > 0 ? 'Bs '+formatNum(ct) : '—'}</td>
        <td>${escHtml(m.cliente||m.proveedor||m.subtipo||'—')}</td>
      </tr>`;
  }).join('');

  const valorInv = inventario
    .filter(p => calcularEstado(p.fecha_vencimiento) !== 'vencido' && parseFloat(p.cantidad) > 0)
    .reduce((s, p) => s + (parseFloat(p.costo_unitario)||0) * parseFloat(p.cantidad), 0);

  const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"/>
<title>${titulo}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#222;padding:24px 32px}
  .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:18px;border-bottom:3px solid #3a6b14;padding-bottom:14px}
  .logo-empresa{font-size:20px;font-weight:900;color:#2d4a18;font-family:Georgia,serif}
  .logo-sub{font-size:10px;color:#666;letter-spacing:2px;text-transform:uppercase;margin-top:2px}
  .consultora-block{display:flex;align-items:center;gap:8px;margin-top:10px;padding:6px 10px;background:#f0f4ff;border:1px solid #c5d0e8;border-radius:4px;width:fit-content}
  .c-name{font-size:12px;font-weight:700;color:#1e3a6e;font-family:Georgia,serif;line-height:1.1}
  .c-sub{font-size:9px;color:#666;letter-spacing:1px;text-transform:uppercase}
  .report-info{text-align:right}
  .report-title{font-size:14px;font-weight:700;color:#333}
  .report-meta{font-size:10px;color:#888;margin-top:3px}
  .summary{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}
  .sum-box{background:#f5f9f0;border:1px solid #cde0b4;border-radius:4px;padding:8px 14px;text-align:center}
  .sum-box strong{display:block;font-size:18px;color:#2d4a18}
  .sum-box span{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:1px}
  .sum-box.warn{background:#fff8e6;border-color:#e0c96e}
  .sum-box.warn strong{color:#7a5200}
  .sum-box.purple{background:#f5f0ff;border-color:#c9a8e8}
  .sum-box.purple strong{color:#6a3db8}
  table{width:100%;border-collapse:collapse}
  thead th{background:#2d4a18;color:#fff;padding:7px 8px;text-align:left;font-size:10px;letter-spacing:.3px}
  tbody tr:nth-child(even){background:#f7faf2}
  tbody td{padding:6px 8px;border-bottom:1px solid #e4edda;vertical-align:middle;font-size:11px}
  .footer{margin-top:18px;font-size:10px;color:#aaa;border-top:1px solid #ddd;padding-top:8px;display:flex;justify-content:space-between}
  @media print{body{padding:0}}
</style></head><body>
<div class="header">
  <div>
    <div class="logo-empresa">🌾 Agropecuaria El Churqui</div>
    <div class="logo-sub">Sistema de Gestión de Inventario</div>
    <div class="consultora-block">
      ${logoB64 ? `<img src="${logoB64}" width="28" height="28" alt="Alto Balance"/>` : ''}
      <div><div class="c-name">Alto Balance</div><div class="c-sub">Consultora</div></div>
    </div>
  </div>
  <div class="report-info">
    <div class="report-title">${titulo}</div>
    <div class="report-meta">Generado: ${new Date().toLocaleString('es-AR')} &nbsp;|&nbsp; ${reporteActual.length} movimientos</div>
  </div>
</div>
<div class="summary">
  <div class="sum-box"><strong>${totalEntradas}</strong><span>Entradas</span></div>
  <div class="sum-box"><strong>${totalSalidas - totalMermas}</strong><span>Ventas</span></div>
  <div class="sum-box warn"><strong>${totalMermas}</strong><span>Mermas</span></div>
  <div class="sum-box"><strong>${reporteActual.length}</strong><span>Total Mov.</span></div>
  ${valorInv > 0 ? `<div class="sum-box purple"><strong>Bs ${formatNum(valorInv)}</strong><span>Valor Inventario</span></div>` : ''}
</div>
<table>
  <thead><tr><th>Fecha</th><th>Hora</th><th>Tipo</th><th>Producto</th><th>Cantidad</th><th>Lote</th><th>Costo Unit.</th><th>Costo Total</th><th>Referencia</th></tr></thead>
  <tbody>${filas}</tbody>
</table>
<div class="footer">
  <span>Agropecuaria El Churqui &nbsp;·&nbsp; Gestionado por Alto Balance Consultora</span>
  <span>Impreso: ${new Date().toLocaleString('es-AR')}</span>
</div>
<script>window.onload=()=>window.print()<\/script>
</body></html>`;

  window.open(URL.createObjectURL(new Blob([html],{type:'text/html'})), '_blank');
  toast('PDF listo para imprimir ✓', 'ok');
}

function formatNum(n) {
  if (!n && n !== 0) return '—';
  return Number(n).toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatBs(n) {
  if (!n && n !== 0) return '—';
  return `Bs ${formatNum(n)}`;
}

function hoy() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function fechaLocalDeISO(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function formatFecha(fecha) {
  if (!fecha) return 'Sin vencimiento';
  try { return new Date(fecha+'T12:00:00').toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit',year:'numeric'}); }
  catch { return fecha; }
}
function formatFecha2(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit',year:'numeric'});
}
function formatHora(iso) {
  if (!iso) return '—';
  try { return new Date(iso).toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit',second:'2-digit'}); }
  catch { return '—'; }
}
function escHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function descargarArchivo(contenido, nombre, tipo) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([contenido],{type:tipo}));
  a.download = nombre;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),5000);
}

// ──────────────────────────────────────────────
// TOAST
// ──────────────────────────────────────────────
let toastTimer;
function toast(msg, tipo = 'ok') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast toast-${tipo} show`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 4000);
}
