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
      <td>${p.fotoNombre || "-"}</td>
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
  const inputFoto = document.getElementById("foto");
  const precioValor = document.getElementById("precio").value;
  const stockValor = document.getElementById("stock").value;
  const descripcion = document.getElementById("descripcion").value.trim();

  if (!nombre) {
    alert("El nombre de la prenda / producto es obligatorio.");
    return null;
  }

  const precio = precioValor !== "" ? Number(precioValor) : null;
  const stock = stockValor !== "" ? Number(stockValor) : null;
  const archivoFoto = inputFoto && inputFoto.files && inputFoto.files[0] ? inputFoto.files[0] : null;
  const fotoNombre = archivoFoto ? archivoFoto.name : "";

  return {
    nombre,
    categoria,
    talla,
    color,
    fotoNombre,
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

  try {
    const wb = XLSX.utils.book_new();

    // Agrupar por categoría para crear una hoja por cada una
    const gruposPorCategoria = new Map();

    productos.forEach((p, idx) => {
      const categoria = p.categoria && p.categoria.trim() ? p.categoria.trim() : "Sin categoría";
      if (!gruposPorCategoria.has(categoria)) {
        gruposPorCategoria.set(categoria, []);
      }
      gruposPorCategoria.get(categoria).push({
        Nº: idx + 1,
        Nombre: p.nombre,
        Categoría: p.categoria,
        Talla: p.talla,
        Color: p.color,
        "Foto (nombre archivo)": p.fotoNombre || "",
        Precio: p.precio,
        Stock: p.stock,
        Descripción: p.descripcion,
      });
    });

    // Crear una hoja por categoría
    gruposPorCategoria.forEach((filas, categoria) => {
      const ws = XLSX.utils.json_to_sheet(filas);

      // Limpiar el nombre de la hoja (Excel no permite algunos caracteres y máximo 31)
      let nombreHoja = categoria.replace(/[:\\/?*\[\]]/g, " ");
      if (!nombreHoja.trim()) nombreHoja = "Productos";
      if (nombreHoja.length > 31) {
        nombreHoja = nombreHoja.substring(0, 31);
      }

      XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
    });

    XLSX.writeFile(wb, "productos_ropa.xlsx");
  } catch (error) {
    console.error("Error al generar el Excel:", error);
    alert("Ocurrió un error al generar el archivo Excel.");
  }
});

// Estado inicial
actualizarEstadoBotones();

