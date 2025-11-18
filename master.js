function cargarWO() {
  return JSON.parse(localStorage.getItem("woData") || "[]");
}

function guardarWO(data) {
  localStorage.setItem("woData", JSON.stringify(data));
}

function actualizarResumen() {
  const data = cargarWO();
  const total = data.length;
  const enEspera = data.filter(w => w.estacion === "En espera").length;
  const finalizadas = data.filter(w => w.estacion === "Finalizada").length;
  const enProceso = total - enEspera - finalizadas;

  document.getElementById("total-wo").textContent = total;
  document.getElementById("en-espera").textContent = enEspera;
  document.getElementById("en-proceso").textContent = enProceso;
  document.getElementById("finalizadas").textContent = finalizadas;
}

function renderizarTabla() {
  const data = cargarWO();
  const tbody = document.querySelector("#woTable tbody");

  tbody.innerHTML = data.map((w, i) => `
    <tr>
      <td>${w.id}</td>
      <td>${w.desc}</td>
      <td>${w.estacion}</td>
      <td>
        ${w.estacion === "En espera" ? 
          `<button onclick="asignarEstacion(${i})">Asignar</button>` : 
          `<button onclick="revertirAWO(${i})">Volver a espera</button>`}
      </td>
    </tr>
  `).join("");

  actualizarResumen();
}

function agregarWO() {
  const id = prompt("Ingrese ID de la WO:");
  if (!id) return;
  const desc = prompt("Ingrese descripción:");
  const data = cargarWO();
  data.push({ id, desc, estacion: "CNC" });
  guardarWO(data);
  renderizarTabla();
}

function asignarEstacion(index) {
  const estaciones = ["CNC", "Molino", "Ensamble", "Lija", "Pintura"];
  const estacion = prompt(`Asignar a estación (${estaciones.join(", ")}):`);
  if (!estacion || !estaciones.includes(estacion)) {
    alert("Estación no válida.");
    return;
  }
  const data = cargarWO();
  data[index].estacion = estacion;
  guardarWO(data);
  renderizarTabla();
}

function revertirAWO(index) {
  const data = cargarWO();
  data[index].estacion = "En espera";
  guardarWO(data);
  renderizarTabla();
}
const nuevaWO = {
  id,
  descripcion,
  piezas,
  clasificacion,
  estado: "En espera", // estación inicial
  historial: [{ estacion: "En espera", fecha: new Date().toLocaleString() }]
};

document.addEventListener("DOMContentLoaded", renderizarTabla);
