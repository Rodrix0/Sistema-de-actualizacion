// Lista en memoria de los productos cargados
let productos = [];

const form = document.getElementById("product-form");
const tbody = document.getElementById("products-tbody");
const btnExport = document.getElementById("btn-export");
const btnClear = document.getElementById("btn-clear");
const emptyMessage = document.getElementById("empty-message");

function actualizarEstadoBotones() {
  const hayProductos = productos.length > 0;
  btnExport.disabled = !hayProductos;
  btnClear.disabled = !hayProductos;
  emptyMessage.style.display = hayProductos ? "none" : "block";
}

function limpiarTabla() {
  tbody.innerHTML = "";
}

function renderTabla() {
  limpiarTabla();

  productos.forEach((p, index) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${p.nombre}</td>
      <td>${p.categoria || "-"}</td>
      <td>${p.talla || "-"}</td>
      <td>${p.color || "-"}</td>
      <td>${p.precio ?? "-"}</td>
      <td>${p.stock ?? "-"}</td>
      <td>${p.descripcion || "-"}</td>
      <td>
        <div class="cell-actions">
          <button type="button" class="btn btn-outline btn-sm" data-index="${index}">Eliminar</button>
        </div>
      </td>
    `;

    tbody.appendChild(tr);
  });

  // Listeners para botones de eliminar
  tbody.querySelectorAll("button[data-index]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = Number(btn.getAttribute("data-index"));
      productos.splice(idx, 1);
      renderTabla();
      actualizarEstadoBotones();
    });
  });
}

function obtenerDatosFormulario() {
  const nombre = document.getElementById("nombre").value.trim();
  const categoria = document.getElementById("categoria").value.trim();
  const talla = document.getElementById("talla").value.trim();
  const color = document.getElementById("color").value.trim();
  const precioValor = document.getElementById("precio").value;
  const stockValor = document.getElementById("stock").value;
  const descripcion = document.getElementById("descripcion").value.trim();

  if (!nombre) {
    alert("El nombre de la prenda / producto es obligatorio.");
    return null;
  }

  const precio = precioValor !== "" ? Number(precioValor) : null;
  const stock = stockValor !== "" ? Number(stockValor) : null;

  return {
    nombre,
    categoria,
    talla,
    color,
    precio,
    stock,
    descripcion,
  };
}

form.addEventListener("submit", (e) => {
  e.preventDefault();
  const datos = obtenerDatosFormulario();
  if (!datos) return;

  productos.push(datos);
  renderTabla();
  actualizarEstadoBotones();

  form.reset();
  document.getElementById("nombre").focus();
});

btnClear.addEventListener("click", () => {
  if (!productos.length) return;
  const confirmar = confirm("¿Seguro que quieres vaciar la lista completa?");
  if (!confirmar) return;
  productos = [];
  renderTabla();
  actualizarEstadoBotones();
});

btnExport.addEventListener("click", () => {
  if (!productos.length) return;

  // Creamos una copia formateada para el Excel, con encabezados más amigables
  const datosExcel = productos.map((p, idx) => ({
    Nº: idx + 1,
    Nombre: p.nombre,
    Categoría: p.categoria,
    Talla: p.talla,
    Color: p.color,
    Precio: p.precio,
    Stock: p.stock,
    Descripción: p.descripcion,
  }));

  try {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(datosExcel);
    XLSX.utils.book_append_sheet(wb, ws, "Productos");
    XLSX.writeFile(wb, "productos_ropa.xlsx");
  } catch (error) {
    console.error("Error al generar el Excel:", error);
    alert("Ocurrió un error al generar el archivo Excel.");
  }
});

// Estado inicial
actualizarEstadoBotones();

