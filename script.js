const APPSCRIPT_URL = "https://script.google.com/macros/s/AKfycbx_-4yuQ2bdDWZ2lzYcedRt_LyU6nzbJx9vNx5EcB39oOneI0Byaa1DkZCxxneBl3DfWg/exec";

let totalItems = 0;

document.getElementById("agregar").addEventListener("click", () => {
  const row = document.createElement("tr");
  row.innerHTML = `
    <td><input type="text" name="sku" required /></td>
    <td><input type="text" name="desc" /></td>
    <td><input type="text" name="tipo" /></td>
    <td><input type="text" name="marca" /></td>
    <td><input type="number" name="cantidad" value="1" min="1" /></td>
    <td><button type="button" onclick="this.closest('tr').remove();updateTotal();">❌</button></td>
  `;
  document.querySelector("#productos-table tbody").appendChild(row);
  updateTotal();
});

document.getElementById("remito-form").addEventListener("submit", async (e) => {
  e.preventDefault();

  const productos = [...document.querySelectorAll("#productos-table tbody tr")].map(row => {
    const cells = row.querySelectorAll("input");
    return {
      sku: cells[0].value,
      descripcion: cells[1].value,
      tipo: cells[2].value,
      marca: cells[3].value,
      cantidad: parseInt(cells[4].value)
    };
  });

  const data = {
    fecha: document.getElementById("fecha").value,
    punto: document.getElementById("punto").value,
    usuario: document.getElementById("usuario").value,
    observaciones: document.getElementById("observaciones").value,
    productos
  };

  try {
    const response = await fetch(APPSCRIPT_URL, {
      method: "POST",
      body: JSON.stringify(data),
      headers: { "Content-Type": "application/json" }
    });
    const result = await response.json();
    alert(result.ok ? "Remito registrado con éxito" : `Error: ${result.error}`);
  } catch (err) {
    alert("Error de red: " + err);
  }
});

function updateTotal() {
  const rows = document.querySelectorAll("#productos-table tbody tr");
  const total = Array.from(rows).reduce((sum, row) => {
    const qty = row.querySelector("input[name='cantidad']");
    return sum + (parseInt(qty?.value) || 0);
  }, 0);
  document.getElementById("total").innerText = total;
}