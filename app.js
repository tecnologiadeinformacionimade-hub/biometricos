<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte Biométrico</title>

<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

<style>
body {
  font-family: Arial, sans-serif;
  background: #f4f6f8;
  padding: 20px;
}

.card {
  background: white;
  padding: 15px;
  border-radius: 8px;
  margin-bottom: 20px;
}

input, select, button {
  padding: 6px 10px;
  margin-right: 10px;
}

button {
  background: #0d6efd;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.export {
  background: #198754;
  float: right;
}

table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 10px;
}

th, td {
  border: 1px solid #ddd;
  padding: 8px;
  text-align: center;
}

th {
  background: #212529;
  color: white;
}
</style>
</head>

<body>

<h2>📊 Reporte Biométrico de Asistencia</h2>

<div class="card">
  <input type="file" id="fileInput" accept=".xlsx,.xls">
  <select id="tipoReporte">
    <option value="SEMANA">Semana</option>
    <option value="MES">Mes</option>
    <option value="ANIO">Año</option>
  </select>
  <input type="date" id="fechaBase">
  <button onclick="procesarExcel()">Generar</button>
</div>

<div class="card">
  <button class="export" onclick="exportarExcel()">Exportar Excel</button>
  <table>
    <thead>
      <tr>
        <th>ID</th>
        <th>Empleado</th>
        <th>Días trabajados</th>
        <th>Horas trabajadas</th>
        <th>Horas de almuerzo</th>
      </tr>
    </thead>
    <tbody id="tablaResultados"></tbody>
  </table>
</div>

<script>
let reporteFinal = [];

// ================================
// UTILIDADES
// ================================
function mismaSemana(fecha, base) {
  const f = new Date(fecha);
  const b = new Date(base);
  const inicio = new Date(b);
  inicio.setDate(b.getDate() - b.getDay());
  const fin = new Date(inicio);
  fin.setDate(inicio.getDate() + 6);
  return f >= inicio && f <= fin;
}

function msToHora(ms) {
  const totalMin = Math.round(ms / 60000);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}

// ================================
// PROCESAR EXCEL
// ================================
function procesarExcel() {
  const file = document.getElementById("fileInput").files[0];
  const tipo = document.getElementById("tipoReporte").value;
  const fechaBase = document.getElementById("fechaBase").value;

  if (!file || !fechaBase) {
    alert("Selecciona archivo y fecha base");
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" });
    generarReporte(data, tipo, new Date(fechaBase));
  };
  reader.readAsBinaryString(file);
}

// ================================
// GENERAR REPORTE (LÓGICA REAL)
// ================================
function generarReporte(data, tipo, base) {
  const empleados = {};
  reporteFinal = [];
  document.getElementById("tablaResultados").innerHTML = "";

  // Agrupar marcajes
  data.forEach(row => {
    const id = Number(row["ID Empleado"]);
    const nombre = row["Nombre"];
    const fecha = row["Fecha"];
    const hora = row["Hora"];

    if (!id || !nombre || !fecha || !hora) return;

    const fh = new Date(`${fecha}T${hora}`);
    if (isNaN(fh)) return;

    let incluir = false;
    if (tipo === "SEMANA") incluir = mismaSemana(fh, base);
    if (tipo === "MES") incluir = fh.getMonth() === base.getMonth() && fh.getFullYear() === base.getFullYear();
    if (tipo === "ANIO") incluir = fh.getFullYear() === base.getFullYear();
    if (!incluir) return;

    empleados[id] ??= { id, nombre, dias: {} };
    empleados[id].dias[fecha] ??= [];
    empleados[id].dias[fecha].push(fh);
  });

  // Calcular totales
  Object.values(empleados).forEach(emp => {
    let dias = 0;
    let trabajoMs = 0;
    let almuerzoMs = 0;

    Object.entries(emp.dias).forEach(([fecha, marcas]) => {
      marcas.sort((a,b) => a-b);
      if (marcas.length < 2) return;

      dias++;

      const entrada = marcas[0];
      const salida = marcas[marcas.length - 1];
      let trabajado = salida - entrada;

      // Almuerzo REAL (marcas 2 y 3)
      let almuerzo = 0;
      if (marcas.length >= 3) {
        const posible = marcas[2] - marcas[1];
        if (posible > 0 && posible <= 3 * 60 * 60 * 1000) {
          almuerzo = posible;
        }
      }

      trabajado -= almuerzo;
      if (trabajado < 0) trabajado = 0;

      trabajoMs += trabajado;
      almuerzoMs += almuerzo;
    });

    reporteFinal.push({
      ID: emp.id,
      Empleado: emp.nombre,
      "Días trabajados": dias,
      "Horas trabajadas": msToHora(trabajoMs),
      "Horas de almuerzo": msToHora(almuerzoMs)
    });
  });

  reporteFinal.sort((a,b) => a.ID - b.ID);

  const tbody = document.getElementById("tablaResultados");
  reporteFinal.forEach(r => {
    tbody.innerHTML += `
      <tr>
        <td>${r.ID}</td>
        <td>${r.Empleado}</td>
        <td>${r["Días trabajados"]}</td>
        <td>${r["Horas trabajadas"]}</td>
        <td>${r["Horas de almuerzo"]}</td>
      </tr>`;
  });
}

// ================================
// EXPORTAR EXCEL
// ================================
function exportarExcel() {
  if (reporteFinal.length === 0) {
    alert("No hay datos para exportar");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(reporteFinal);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Resumen");
  XLSX.writeFile(wb, "reporte_biometrico_final.xlsx");
}
</script>

</body>
</html>
