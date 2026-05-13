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

    // Función auxiliar para procesar y formatear una hoja
    const crearHoja = (datosFiltrados, nombreHojaBase) => {
      // 1. Preparar los datos con nombres de columna amigables
      const filas = datosFiltrados.map((p, idx) => ({
        "Nº": idx + 1,
        "Nombre": p.nombre,
        "Categoría": p.categoria || "Sin categoría",
        "Talla": p.talla || "-",
        "Color": p.color || "-",
        "Foto": p.fotoNombre || "-",
        "Precio": p.precio || 0,
        "Stock": p.stock || 0,
        "Descripción": p.descripcion || "-",
      }));

      // 2. Crear la hoja
      const ws = XLSX.utils.json_to_sheet(filas);

      // 3. Calcular totales
      const totalPrecio = datosFiltrados.reduce((sum, p) => sum + (p.precio || 0), 0);
      const totalStock = datosFiltrados.reduce((sum, p) => sum + (p.stock || 0), 0);

      // 4. Añadir fila de totales al final
      const ultimaFila = filas.length + 1;
      const filaTotales = [
        ["", "", "", "", "TOTALES", "", totalPrecio, totalStock, ""]
      ];
      XLSX.utils.sheet_add_aoa(ws, filaTotales, { origin: `A${ultimaFila + 1}` });

      // 5. Configurar formatos (Moneda para Precio en columna G)
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let R = 1; R <= range.e.r; ++R) {
        const cellPrecio = ws[XLSX.utils.encode_cell({ r: R, c: 6 })]; // Columna G (index 6)
        if (cellPrecio && typeof cellPrecio.v === 'number') {
          cellPrecio.z = '"$"#,##0.00'; // Formato de moneda
        }
      }

      // 6. Configurar anchos de columna (ajuste manual aproximado)
      ws['!cols'] = [
        { wch: 5 },  // Nº
        { wch: 25 }, // Nombre
        { wch: 15 }, // Categoría
        { wch: 10 }, // Talla
        { wch: 12 }, // Color
        { wch: 20 }, // Foto
        { wch: 12 }, // Precio
        { wch: 10 }, // Stock
        { wch: 30 }, // Descripción
      ];

      // 7. Limpiar y validar nombre de hoja
      let nombreHoja = nombreHojaBase.replace(/[:\\/?*\[\]]/g, " ");
      if (!nombreHoja.trim()) nombreHoja = "Hoja";
      nombreHoja = nombreHoja.substring(0, 31);

      XLSX.utils.book_append_sheet(wb, ws, nombreHoja);
    };

    // --- Generar Hoja General ---
    crearHoja(productos, "General");

    // --- Generar Hojas por Categoría ---
    const categorias = [...new Set(productos.map(p => p.categoria && p.categoria.trim() ? p.categoria.trim() : "Sin categoría"))];
    
    // Si solo hay una categoría, no es necesario duplicar (opcional, pero mejor ser explícito)
    categorias.forEach(cat => {
      const productosCat = productos.filter(p => (p.categoria?.trim() || "Sin categoría") === cat);
      crearHoja(productosCat, cat);
    });

    // Guardar archivo
    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Inventario_Productos_${fecha}.xlsx`);

  } catch (error) {
    console.error("Error al generar el Excel:", error);
    alert("Ocurrió un error al generar el archivo Excel.");
  }
});

// Estado inicial
actualizarEstadoBotones();

