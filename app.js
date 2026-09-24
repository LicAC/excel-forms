// =====================================================================================
// FORMULARIO DE REGISTRO DE ACTIVIDADES - ADD-IN SCRIPT LAB PARA EXCEL
// Estilo inspirado en la vista CRM / Gestión Comercial (ComerciaSync)
// =====================================================================================

//#region 1. INTERFACES Y TIPOS

//#endregion

//#region 2. INICIALIZACIÓN
Office.onReady(() => {
  // Escuchar el evento click del botón Guardar
  const btnGuardar = document.getElementById("btnGuardar");
  if (btnGuardar) {
    btnGuardar.addEventListener("click", e => tryCatch(() => procesarFormulario(e)));
  }
});

/** Helper por defecto de Script Lab para manejar errores */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
//#endregion

//#region 3. PROCESAMIENTO Y VALIDACIÓN
async function procesarFormulario(evento) {
  // Prevenir cualquier comportamiento por defecto si aplica
  if (evento) evento.preventDefault();

  // Obtener referencias a los elementos HTML
  const inputTitulo = document.getElementById("titulo");
  const inputEmpresa = document.getElementById("empresa");
  const inputMonto = document.getElementById("monto");
  const selectEstado = document.getElementById("estado");
  const radioPrioridadChecked = document.querySelector('input[name="prioridad"]:checked');

  // Limpiar estilos de validación previos
  [inputTitulo, inputEmpresa, inputMonto, selectEstado].forEach(el => {
    if (el) el.classList.remove("invalid");
  });

  // Extraer valores limpiando espacios en blanco
  const titulo = inputTitulo ? inputTitulo.value.trim() : "";
  const empresa = inputEmpresa ? inputEmpresa.value.trim() : "";
  const monto = inputMonto ? Number(inputMonto.value) : NaN;
  const estado = selectEstado ? selectEstado.value : "Nuevo";
  const prioridad = radioPrioridadChecked ? radioPrioridadChecked.value : "Media";

  // Validar campos obligatorios
  let esValido = true;
  if (!titulo) {
    inputTitulo.classList.add("invalid");
    esValido = false;
  }
  if (!empresa) {
    inputEmpresa.classList.add("invalid");
    esValido = false;
  }
  if (!inputMonto || !inputMonto.value || isNaN(monto) || monto < 0) {
    inputMonto.classList.add("invalid");
    esValido = false;
  }

  // Si no es válido, notificar al usuario y detener la ejecución
  if (!esValido) {
    mostrarToast("⚠️ Completa los campos marcados en rojo.", "warning");
    return;
  }
  try {
    // 1. Obtener el siguiente ID autoincremental desde Excel
    const nuevoId = await obtenerSiguienteId();

    // 2. Crear el objeto con la nueva actividad
    const nuevaActividad = {
      id: nuevoId,
      titulo: titulo,
      empresa: empresa,
      monto: monto,
      prioridad: prioridad,
      estado: estado
    };

    // 3. Guardar la actividad en la tabla de Excel
    await guardarActividadEnExcel(nuevaActividad);

    // 4. Feedback visual de éxito (Toast)
    mostrarToast(`✅ ¡Actividad #${nuevoId} registrada con éxito!`, "success");

    // 5. Limpiar el formulario y resetear selección
    limpiarFormulario();
  } catch (error) {
    console.error(error);
    mostrarToast("❌ Ocurrió un error al guardar en Excel.", "error");
  }
}
//#endregion

//#region 4. OPERACIONES EN EXCEL (API DE OFFICE)
async function obtenerSiguienteId() {
  // Consulta la tabla 'tblKanban' y retorna el máximo ID + 1
  let siguienteId = 1;
  await Excel.run(async context => {
    const table = context.workbook.tables.getItem("tblKanban");
    const bodyRange = table.getDataBodyRange();

    // Cargar sólo los datos de la tabla
    bodyRange.load("values");
    await context.sync();
    const filas = bodyRange.values;

    // Si existen filas, encontrar el ID máximo en la primera columna (índice 0)
    if (filas && filas.length > 0) {
      const ids = filas.map(row => Number(row[0])).filter(id => !isNaN(id));
      if (ids.length > 0) {
        siguienteId = Math.max(...ids) + 1;
      }
    }
  });
  return siguienteId;
}
async function guardarActividadEnExcel(actividad) {
  // Agrega una nueva fila al final de la tabla 'tblKanban'
  await Excel.run(async context => {
    const table = context.workbook.tables.getItem("tblKanban");

    // Formatear la fila en el mismo orden que las columnas de Excel
    // ID, TITULO, EMPRESA, MONTO, PRIORIDAD, ESTADO
    const nuevaFila = [[actividad.id, actividad.titulo, actividad.empresa, actividad.monto, actividad.prioridad, actividad.estado]];

    // Insertar la fila en la tabla de Excel
    table.rows.add(null, nuevaFila);
    await context.sync();
  });
}
//#endregion

//#region 5. UTILIDADES DE INTERFAZ (TOAST Y LIMPIEZA)
function mostrarToast(mensaje, tipo) {
  const container = document.getElementById("toastContainer");
  if (!container) return;

  // Crear el elemento de notificación
  const toast = document.createElement("div");
  toast.className = `toast ${tipo}`;

  // Agregar icono según el tipo
  let icono = "check_circle";
  if (tipo === "warning") icono = "warning";
  if (tipo === "error") icono = "error";
  toast.innerHTML = `<span class="material-symbols-outlined">${icono}</span><span>${mensaje}</span>`;
  container.appendChild(toast);

  // Remover automáticamente después de 3.5 segundos
  setTimeout(() => {
    toast.remove();
  }, 3500);
}
function limpiarFormulario() {
  const inputTitulo = document.getElementById("titulo");
  const inputEmpresa = document.getElementById("empresa");
  const inputMonto = document.getElementById("monto");
  const selectEstado = document.getElementById("estado");
  if (inputTitulo) inputTitulo.value = "";
  if (inputEmpresa) inputEmpresa.value = "";
  if (inputMonto) inputMonto.value = "";
  if (selectEstado) selectEstado.value = "Nuevo";

  // Regresar la prioridad por defecto a 'Media'
  const radioMedia = document.querySelector('input[name="prioridad"][value="Media"]');
  if (radioMedia) {
    radioMedia.checked = true;
  }

  // Foco en el primer campo
  if (inputTitulo) {
    inputTitulo.focus();
  }
}
//#endregion