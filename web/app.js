// Panel de Gastos 2026 — lógica del frontend
"use strict";

let DATOS = null;
let mesSeleccionado = null; // nombre de hoja seleccionado para el filtro de mes
let ultimoMtime = null;

const $ = (id) => document.getElementById(id);

// Si se abre como archivo local (file://) en vez de a través del servidor
if (location.protocol === "file:") {
  $("excel-info").textContent = "⚠️ Abre http://127.0.0.1:8720 (ejecuta iniciar.bat)";
  $("excel-info").className = "excel-info err";
  mostrarAviso("Esta página debe abrirse a través del servidor. Ejecuta iniciar.bat y abre http://127.0.0.1:8720", "error");
}

function fmtEuro(n) {
  if (n === null || n === undefined) return "—";
  return n.toLocaleString("es-ES", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }) + " €";
}

function signo(n) {
  if (n === null || n === undefined) return "";
  return n < 0 ? "neg" : "pos";
}

function formatoValor(valor, tipo) {
  if (valor === null || valor === undefined) return "—";
  if (tipo === "euro") return fmtEuro(valor);
  return String(valor);
}

function toast(mensaje, tipo = "") {
  const el = $("toast");
  el.textContent = mensaje;
  el.className = "toast " + tipo;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.add("oculto"), 3200);
}

function mostrarAviso(mensaje, tipo = "error") {
  const el = $("aviso");
  if (!mensaje) {
    el.classList.add("oculto");
    return;
  }
  el.textContent = mensaje;
  el.className = "aviso " + tipo;
}

// ------------------------------------------------------------
// Carga de datos
// ------------------------------------------------------------

async function cargar() {
  mostrarAviso("");
  try {
    const resp = await fetch("/api/datos");
    const datos = await resp.json();
    if (datos.error) {
      DATOS = null;
      $("excel-info").textContent = datos.error;
      $("excel-info").className = "excel-info err";
      mostrarAviso(datos.error, "error");
      limpiarTodo();
      return;
    }
    DATOS = datos;
    ultimoMtime = datos.excel_mtime;
    $("excel-info").textContent =
      "📁 " + datos.excel_nombre + " · actualizado " + datos.excel_mtime;
    $("excel-info").className = "excel-info ok";
    render();
    actualizarIndicador();
  } catch (error) {
    DATOS = null;
    $("excel-info").textContent = "No se pudo conectar con el servidor.";
    $("excel-info").className = "excel-info err";
    mostrarAviso("No se pudo conectar con el servidor local. ¿Está ejecutándose servidor.py?", "error");
    limpiarTodo();
  }
}

function limpiarTodo() {
  $("resumen-anual").innerHTML = "";
  $("chart-comparativa").innerHTML = "";
  $("grid-meses").innerHTML = "";
  document.querySelector("#tabla-gastos tbody").innerHTML = "";
  $("resumen-filtrado").textContent = "";
}

// ------------------------------------------------------------
// Render
// ------------------------------------------------------------

function render() {
  renderResumenAnual();
  renderComparativa();
  renderTablaMensual();
  renderMeses();
  renderFiltros();
  renderTabla();
}

function renderResumenAnual() {
  const a = DATOS.anual;
  const totalGastosRegistrados = DATOS.meses.reduce((s, m) => s + (m.n_expenses || 0), 0);
  const items = [
    { etiqueta: "Gastos totales (año)", valor: a.gastos_total, tipo: "euro" },
    { etiqueta: "Disponible total (año)", valor: a.disponible_total, tipo: "euro" },
    { etiqueta: "Ingresos totales (año)", valor: a.ingresos_total, tipo: "euro" },
    { etiqueta: "Ahorro total del año", valor: a.ahorro_total, tipo: "euro", clase: signo(a.ahorro_total) },
    { etiqueta: "Gastos registrados", valor: totalGastosRegistrados, tipo: "num" },
    { etiqueta: "Meses con datos", valor: a.n_meses, tipo: "num" },
  ];
  $("resumen-anual").innerHTML = items.map((it) => `
    <div class="tarjeta">
      <div class="etiqueta">${it.etiqueta}</div>
      <div class="valor ${it.clase || ""}">${formatoValor(it.valor, it.tipo)}</div>
    </div>
  `).join("");
}

function renderComparativa() {
  const meses = DATOS.meses;
  const max = Math.max(1, ...meses.map((m) => Math.abs(m.gastos_total || 0)));
  $("chart-comparativa").innerHTML = meses.map((m) => {
    const valor = m.gastos_total || 0;
    const pct = max ? (Math.abs(valor) / max) * 100 : 0;
    const neg = valor < 0;
    return `
      <div class="chart-fila">
        <div class="chart-mes">${m.mes.toLowerCase()}</div>
        <div class="chart-track">
          <div class="chart-barra ${neg ? "neg" : ""}" style="width:${pct}%"></div>
        </div>
        <div class="chart-valor">${fmtEuro(valor)}</div>
      </div>
    `;
  }).join("");
}

function renderTablaMensual() {
  const filas = DATOS.meses.map((m) => {
    const ingresos = (m.ingresos === null || m.ingresos === undefined) ? null : m.ingresos;
    const gastos = (m.gastos_total === null || m.gastos_total === undefined) ? null : m.gastos_total;
    const otro = (m.otro_gasto === null || m.otro_gasto === undefined) ? null : m.otro_gasto;
    const dif = (ingresos !== null && gastos !== null) ? Math.round((ingresos - gastos) * 100) / 100 : null;
    const saldo = (dif !== null && otro !== null) ? Math.round((dif - otro) * 100) / 100 : null;
    return { mes: m.mes, anio: m.anio, ingresos, gastos, otro, dif, saldo };
  });

  document.querySelector("#tabla-mensual tbody").innerHTML = filas.map((f) => `
    <tr>
      <td class="mes-etiqueta">${escapeHtml(f.mes.toLowerCase())} ${f.anio}</td>
      <td class="num">${fmtEuro(f.ingresos)}</td>
      <td class="num">${fmtEuro(f.gastos)}</td>
      <td class="num ${signo(f.dif)}">${fmtEuro(f.dif)}</td>
      <td class="num">${fmtEuro(f.otro)}</td>
      <td class="num ${signo(f.saldo)}">${fmtEuro(f.saldo)}</td>
    </tr>
  `).join("");

  const totIng = Math.round(filas.reduce((s, f) => s + (f.ingresos || 0), 0) * 100) / 100;
  const totGas = Math.round(filas.reduce((s, f) => s + (f.gastos || 0), 0) * 100) / 100;
  const totOtro = Math.round(filas.reduce((s, f) => s + (f.otro || 0), 0) * 100) / 100;
  const totDif = Math.round((totIng - totGas) * 100) / 100;
  const totSaldo = Math.round((totDif - totOtro) * 100) / 100;
  document.querySelector("#tabla-mensual tfoot").innerHTML = `
    <tr class="fila-total">
      <td>TOTAL</td>
      <td class="num">${fmtEuro(totIng)}</td>
      <td class="num">${fmtEuro(totGas)}</td>
      <td class="num ${signo(totDif)}">${fmtEuro(totDif)}</td>
      <td class="num">${fmtEuro(totOtro)}</td>
      <td class="num ${signo(totSaldo)}">${fmtEuro(totSaldo)}</td>
    </tr>
  `;
}

function renderMeses() {
  $("grid-meses").innerHTML = DATOS.meses.map((m) => {
    const activa = mesSeleccionado === m.nombre_hoja ? "activa" : "";
    return `
      <div class="mes-card ${activa}" data-hoja="${escapeHtml(m.nombre_hoja)}">
        <h3>${m.mes.toLowerCase()} ${m.anio}</h3>
        <div class="fila"><span>Gastos</span><span>${fmtEuro(m.gastos_total)}</span></div>
        <div class="fila"><span>Disponible</span><span>${fmtEuro(m.disponible)}</span></div>
        <div class="fila"><span>Diferencia</span><span class="${signo(m.diferencia)}">${fmtEuro(m.diferencia)}</span></div>
        <div class="fila"><span>Ahorro final</span><span class="${signo(m.ahorro)}">${fmtEuro(m.ahorro)}</span></div>
        <div class="fila"><span>Nº gastos</span><span>${m.n_expenses}</span></div>
      </div>
    `;
  }).join("");

  document.querySelectorAll(".mes-card").forEach((card) => {
    card.addEventListener("click", () => {
      const hoja = card.getAttribute("data-hoja");
      mesSeleccionado = (mesSeleccionado === hoja) ? null : hoja;
      $("filtro-mes").value = mesSeleccionado || "";
      renderMeses();
      renderTabla();
    });
  });
}

function renderFiltros() {
  const catSel = $("filtro-categoria");
  const mesSel = $("filtro-mes");
  const prevCat = catSel.value;
  const prevMes = mesSel.value;

  catSel.innerHTML = '<option value="">Todas las categorías</option>' +
    DATOS.categorias_disponibles.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  mesSel.innerHTML = '<option value="">Todos los meses</option>' +
    DATOS.meses.map((m) => `<option value="${escapeHtml(m.nombre_hoja)}">${escapeHtml(m.mes.toLowerCase())} ${m.anio}</option>`).join("");

  catSel.value = prevCat;
  mesSel.value = prevMes;
  if (!mesSeleccionado && prevMes) mesSeleccionado = prevMes;

  // Datalist de categorías para el formulario
  $("lista-categorias").innerHTML = DATOS.categorias_disponibles
    .map((c) => `<option value="${escapeHtml(c)}"></option>`).join("");
}

function renderTabla() {
  const texto = $("buscar").value.trim().toLowerCase();
  const cat = $("filtro-categoria").value;
  const hoja = $("filtro-mes").value;

  const filas = [];
  for (const m of DATOS.meses) {
    if (hoja && m.nombre_hoja !== hoja) continue;
    for (const g of m.gastos) {
      if (cat && g.categoria !== cat) continue;
      if (texto && !g.concepto.toLowerCase().includes(texto)) continue;
      filas.push({ ...g, mes: m.mes.toLowerCase(), anio: m.anio });
    }
  }

  // orden: más recientes primero (año/mes/día descendente por hoja)
  filas.sort((a, b) => {
    if (a.anio !== b.anio) return b.anio - a.anio;
    const mi = DATOS.meses.find((m) => m.nombre_hoja === a._hoja);
    return (b.dia || 0) - (a.dia || 0);
  });

  // orden estable por mes (hoja) y día
  filas.sort((a, b) => {
    const idxA = DATOS.meses.findIndex((m) => m.mes.toLowerCase() === a.mes);
    const idxB = DATOS.meses.findIndex((m) => m.mes.toLowerCase() === b.mes);
    if (idxA !== idxB) return idxB - idxA;
    return (b.dia || 0) - (a.dia || 0);
  });

  const total = filas.reduce((s, f) => s + f.importe, 0);
  $("resumen-filtrado").textContent =
    `${filas.length} gasto(s) · total ${fmtEuro(total)}`;

  const tbody = document.querySelector("#tabla-gastos tbody");
  tbody.innerHTML = filas.map((f) => `
    <tr>
      <td><span class="mes-etiqueta">${escapeHtml(f.mes)}</span></td>
      <td>${f.dia || "—"}</td>
      <td>${escapeHtml(f.concepto)}</td>
      <td><span class="chip">${escapeHtml(f.categoria)}</span></td>
      <td class="num">${fmtEuro(f.importe)}</td>
    </tr>
  `).join("");

  // Fila de TOTAL al final de la lista filtrada
  document.querySelector("#tfoot-gastos").innerHTML = `
    <tr class="fila-total">
      <td colspan="4">TOTAL · ${filas.length} gasto(s)</td>
      <td class="num">${fmtEuro(total)}</td>
    </tr>
  `;
}

// ------------------------------------------------------------
// Filtros (eventos)
// ------------------------------------------------------------

$("buscar").addEventListener("input", renderTabla);
$("filtro-categoria").addEventListener("change", renderTabla);
$("filtro-mes").addEventListener("change", (e) => {
  mesSeleccionado = e.target.value || null;
  renderMeses();
  renderTabla();
});

// ------------------------------------------------------------
// Modal añadir gasto
// ------------------------------------------------------------

function abrirModal() {
  $("modal").classList.remove("oculto");
  $("form-error").classList.add("oculto");
  const hoy = new Date();
  $("fecha").value = hoy.toISOString().slice(0, 10);
  $("concepto").value = "";
  $("importe").value = "";
  $("categoria").value = "";
  $("concepto").focus();
}

function cerrarModal() {
  $("modal").classList.add("oculto");
}

$("btn-nuevo").addEventListener("click", abrirModal);
$("btn-cancelar").addEventListener("click", cerrarModal);
$("btn-cerrar-modal").addEventListener("click", cerrarModal);
$("modal").addEventListener("click", (e) => {
  if (e.target === $("modal")) cerrarModal();
});

$("form-gasto").addEventListener("submit", async (e) => {
  e.preventDefault();
  $("form-error").classList.add("oculto");

  const payload = {
    fecha: $("fecha").value,
    concepto: $("concepto").value.trim(),
    importe: $("importe").value,
    categoria: $("categoria").value.trim(),
  };

  if (!payload.fecha || !payload.concepto || !payload.importe) {
    $("form-error").textContent = "Rellena fecha, concepto e importe.";
    $("form-error").classList.remove("oculto");
    return;
  }

  const btn = document.querySelector("#form-gasto button[type=submit]");
  btn.disabled = true;
  btn.textContent = "Guardando…";

  try {
    const resp = await fetch("/api/gasto", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const datos = await resp.json();
    if (datos.ok) {
      cerrarModal();
      toast("✅ Gasto añadido: " + datos.resultado.concepto + " · " +
            fmtEuro(datos.resultado.importe), "ok");
      await cargar();
    } else {
      $("form-error").textContent = datos.error || "Error desconocido.";
      $("form-error").classList.remove("oculto");
    }
  } catch (error) {
    $("form-error").textContent = "No se pudo conectar con el servidor.";
    $("form-error").classList.remove("oculto");
  } finally {
    btn.disabled = false;
    btn.textContent = "Guardar";
  }
});

// ------------------------------------------------------------
// Utilidades
// ------------------------------------------------------------

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

// ------------------------------------------------------------
// Inicio
// ------------------------------------------------------------

$("btn-actualizar").addEventListener("click", cargar);

// ------------------------------------------------------------
// Auto-actualización (comprueba el Excel cada 10 segundos)
// ------------------------------------------------------------

function actualizarIndicador() {
  const el = $("auto-status");
  if (!el) return;
  const hora = new Date().toLocaleTimeString("es-ES");
  el.textContent = "🟢 Auto cada 10s · " + hora;
}

async function refrescarSilencioso() {
  if (!DATOS || document.hidden) return;
  try {
    const resp = await fetch("/api/datos");
    const datos = await resp.json();
    if (datos.error || !datos.excel_mtime) return;
    if (datos.excel_mtime !== ultimoMtime) {
      ultimoMtime = datos.excel_mtime;
      DATOS = datos;
      render();
      actualizarIndicador();
    }
  } catch (e) {
    // servidor caído; se reintenta en la siguiente pasada
  }
}

setInterval(refrescarSilencioso, 10000);
cargar();
