// === FUNCIONES BASE ===

// ---------- UTILIDADES (coloca al inicio o donde tengas similares) ----------
function obtenerWO() {
  return JSON.parse(localStorage.getItem("workOrders")) || [];
}
function guardarWO(data) {
  localStorage.setItem("workOrders", JSON.stringify(data));
}

// ---------- FUNCION PARA IMPORTAR UNA FILA DESDE EXCEL ----------
/*
  recibe un objeto con campos: { id, descripcion, piezas, clasificacion, estado, historial }
  - si id existe: actualiza la WO existente (merge)
  - si id no existe: genera un id nuevo (WO-001 style) y la agrega
  devuelve: 'added' o 'updated'
*/
window.__importWO = function (row) {
  const all = obtenerWO();

  // Normalizar id
  let id = row.id ? String(row.id).trim() : null;

  // Si no trae id, generar uno secuencial tipo WO-001
  if (!id) {
    // buscar último id con prefijo WO-###
    let lastNum = 0;
    all.forEach(w => {
      if (typeof w.id === 'string' && w.id.startsWith('WO-')) {
        const n = parseInt(w.id.split('-')[1]) || 0;
        if (n > lastNum) lastNum = n;
      }
    });
    id = `WO-${String(lastNum + 1).padStart(3, '0')}`;
  }

  // Buscar si ya existe
  const idx = all.findIndex(w => String(w.id) === String(id));
  if (idx >= 0) {
    // actualizar: merge de campos (prioriza los del Excel si no vacíos)
    const existing = all[idx];
    existing.descripcion = row.descripcion || existing.descripcion;
    existing.piezas = (row.piezas !== undefined && row.piezas !== '') ? Number(row.piezas) : existing.piezas;
    existing.clasificacion = row.clasificacion || existing.clasificacion;
    existing.estado = row.estado || existing.estado || 'En espera';
    // merge historial: si viene historial usamos; si no, preservamos
    if (Array.isArray(row.historial) && row.historial.length) {
      existing.historial = row.historial;
    } else if (!existing.historial) {
      existing.historial = [{ estacion: existing.estado, fecha: new Date().toLocaleString() }];
    }
    all[idx] = existing;
    guardarWO(all);
    return 'updated';
  } else {
    // crear nueva
    const nueva = {
      id,
      descripcion: row.descripcion || '',
      piezas: Number(row.piezas) || 0,
      clasificacion: row.clasificacion || '',
      estado: row.estado || 'En espera',
      historial: Array.isArray(row.historial) && row.historial.length ? row.historial : [{ estacion: row.estado || 'En espera', fecha: new Date().toLocaleString() }]
    };
    all.push(nueva);
    guardarWO(all);
    return 'added';
  }
};




function obtenerWO() {
  return JSON.parse(localStorage.getItem("workOrders")) || [];
}

function guardarWO(data) {
  localStorage.setItem("workOrders", JSON.stringify(data));
}

// === ACTUALIZA TODAS LAS TABLAS DE LAS ESTACIONES ===
function actualizarTablas() {
  actualizarTablaPorEstacion("CNC");
  actualizarTablaPorEstacion("Molino");
  actualizarTablaPorEstacion("Ensamble");
  actualizarTablaPorEstacion("Empaque");
  actualizarTablaMaster();
}

// === AVANZAR WO ENTRE ESTACIONES ===
function avanzarWO(id, nuevaEstacion) {
  const allWOs = obtenerWO();
  const wo = allWOs.find(w => w.id === id);
  if (!wo) return alert("WO no encontrada");

  wo.estado = nuevaEstacion;
  wo.historial = wo.historial || [];
  wo.historial.push({
    estacion: nuevaEstacion,
    fecha: new Date().toLocaleString()
  });

  guardarWO(allWOs);
  actualizarTablas(); // refresca todas las vistas
}

// === RENDERIZA TABLAS POR ESTACIÓN ===
function actualizarTablaPorEstacion(estacion) {
  const tbody = document.getElementById(`tabla-${estacion.toLowerCase()}`);
  if (!tbody) return;

  const allWOs = obtenerWO();
  tbody.innerHTML = "";

  allWOs
    .filter(wo => wo.estado === estacion)
    .forEach(wo => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${wo.id}</td>
        <td>${wo.descripcion}</td>
        <td>${wo.piezas}</td>
        <td>${wo.clasificacion}</td>
        <td>${wo.estado}</td>
        <td>
          <select onchange="avanzarWO('${wo.id}', this.value)">
            <option value="">Mover a...</option>
            <option value="CNC">CNC</option>
            <option value="Molino">Molino</option>
            <option value="Ensamble">Ensamble</option>
            <option value="Empaque">Empaque</option>
            <option value="Finalizado">Finalizado</option>
          </select>
        </td>
      `;
      tbody.appendChild(tr);
    });
}

// === MASTER ===
function actualizarTablaMaster() {
  const tbody = document.getElementById("tabla-master");
  if (!tbody) return;

  const allWOs = obtenerWO();
  tbody.innerHTML = "";

  allWOs.forEach(wo => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${wo.id}</td>
      <td>${wo.descripcion}</td>
      <td>${wo.piezas}</td>
      <td>${wo.clasificacion}</td>
      <td>${wo.estado}</td>
      <td>
        <select onchange="avanzarWO('${wo.id}', this.value)">
          <option value="">Mover a...</option>
          <option value="CNC">CNC</option>
          <option value="Molino">Molino</option>
          <option value="Ensamble">Ensamble</option>
          <option value="Empaque">Empaque</option>
          <option value="Finalizado">Finalizado</option>
        </select>
      </td>
    `;
    tbody.appendChild(tr);
  });
}
// === DASHBOARD GENERAL (index.html) ===
// === DASHBOARD GENERAL (index.html) ===
function actualizarTablaDashboard() {
  const tbody = document.getElementById("tabla-dashboard");
  if (!tbody) return;

  const allWOs = JSON.parse(localStorage.getItem("workOrders")) || [];
  tbody.innerHTML = "";

  allWOs.forEach(wo => {
    const tr = document.createElement("tr");

    // Determinar clase de color según la estación
    let claseColor = "";
    switch (wo.estado) {
      case "En espera": claseColor = "estado-espera"; break;
      case "CNC": claseColor = "estado-cnc"; break;
      case "Molino": claseColor = "estado-molino"; break;
      case "Ensamble": claseColor = "estado-ensamble"; break;
      case "Lija": claseColor = "estado-lija"; break;
      case "Pintura": claseColor = "estado-pintura"; break;
      case "Finalizado": claseColor = "estado-finalizado"; break;
    }

    tr.innerHTML = `
      <td>${wo.id}</td>
      <td>${wo.descripcion}</td>
      <td>${wo.piezas}</td>
      <td>${wo.clasificacion}</td>
      <td class="${claseColor}">${wo.estado}</td>
    `;

    tbody.appendChild(tr);
  });
}

function agregarWO(descripcion, piezas, clasificacion) {
  const allWOs = JSON.parse(localStorage.getItem("workOrders")) || [];
  const nuevaWO = {
    id: Date.now().toString(),
    descripcion,
    piezas,
    clasificacion,
    estado: "En espera"
  };
  allWOs.push(nuevaWO);
  localStorage.setItem("workOrders", JSON.stringify(allWOs));
  actualizarTablas(); // o la función que refresca tus tablas
}

// 🔄 Actualizar automáticamente desde SharePoint en todas las terminales
async function actualizarDesdeSharePoint() {
  try {
    const response = await fetch("https://grieder-my.sharepoint.com/personal/santiago_almazan_glennrieder_com/Documents/Plantilla_WOs.xlsx");
    if (!response.ok) throw new Error("No se pudo acceder al archivo Excel");

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const ultimaVersionLocal = JSON.parse(localStorage.getItem("workOrders")) || [];

    // Solo actualizar si hay cambios
    const jsonLocal = JSON.stringify(ultimaVersionLocal);
    const jsonRemoto = JSON.stringify(rows);

    if (jsonLocal !== jsonRemoto) {
      localStorage.setItem("workOrders", JSON.stringify(rows));
      mostrarTablaMaster();
      console.log("✅ Datos actualizados desde SharePoint");
    } else {
      console.log("ℹ️ Datos ya estaban sincronizados");
    }

  } catch (err) {
    console.error("❌ Error al actualizar desde SharePoint:", err);
  }
}
async function guardarCambiosEnSharePoint() {
  const allWOs = JSON.parse(localStorage.getItem("workOrders")) || [];
  const ws = XLSX.utils.json_to_sheet(allWOs);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "WOs");

  // Convierte a Blob
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  // 🔧 Debes colocar aquí el enlace directo al archivo en SharePoint
  // Este enlace debe tener permisos de edición para que se pueda sobrescribir
  const sharepointUploadUrl = "https://grieder-my.sharepoint.com/personal/santiago_almazan_glennrieder_com/Documents/Plantilla_WOs.xlsx";

  try {
    const response = await fetch(sharepointUploadUrl, {
      method: "PUT",
      body: blob,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    });

    if (response.ok) {
      console.log("✅ Cambios guardados en SharePoint");
    } else {
      console.error("❌ Error al subir a SharePoint:", response.statusText);
    }
  } catch (error) {
    console.error("⚠️ No se pudo conectar con SharePoint:", error);
  }
}

// 🔁 Revisa automáticamente cada 60 segundos si hay una versión nueva
setInterval(actualizarDesdeSharePoint, 60000);

// ⚡ Ejecutar también al cargar la página
document.addEventListener("DOMContentLoaded", actualizarDesdeSharePoint);


document.addEventListener("DOMContentLoaded", actualizarTablas);

