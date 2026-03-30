let columns = [
  "No",
  "No Resi",
  "Customer",
  "Origin",
  "Destination",
  "Tanggal Scan In",
  "POD Digital",
  "Status AP",
  "SLA (Hari)",
  "Aging",
  "Last Update"
];

let data = [
  [1, "JNE001239812", "PT Maju Jaya", "Jakarta", "Bandung", "2026-03-01", "Yes", "AP Verified", 2, 1, "2026-03-02"],
  [2, "JNE001239813", "CV Sukses Selalu", "Surabaya", "Jakarta", "2026-03-02", "Yes", "AP Verified", 3, 2, "2026-03-04"],
  [3, "JNE001239814", "PT Abadi Sentosa", "Medan", "Jakarta", "2026-03-01", "No", "Pending", 4, 5, "2026-03-06"],
  [4, "JNE001239815", "PT Makmur Sejahtera", "Jakarta", "Semarang", "2026-03-03", "Yes", "AP Verified", 2, 1, "2026-03-04"],
  [5, "JNE001239816", "CV Berkah", "Bandung", "Surabaya", "2026-03-01", "No", "Pending", 3, 6, "2026-03-07"],
  [6, "JNE001239817", "PT Nusantara", "Jakarta", "Bali", "2026-03-02", "Yes", "AP Verified", 5, 3, "2026-03-05"],
  [7, "JNE001239818", "PT Global Tech", "Batam", "Jakarta", "2026-03-02", "Yes", "AP Verified", 2, 2, "2026-03-04"],
  [8, "JNE001239819", "PT Logistik Indo", "Makassar", "Jakarta", "2026-03-01", "No", "Pending", 6, 7, "2026-03-08"]
];

function renderTable() {
  const thead = document.getElementById("thead");
  const tbody = document.getElementById("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  // Header
  let header = "<tr>";
  columns.forEach(col => {
    header += `<th>${col}</th>`;
  });
  header += "</tr>";
  thead.innerHTML = header;

  // Body
  data.forEach((row, rIndex) => {
    let tr = "<tr>";
    row.forEach((cell, cIndex) => {
      tr += `
        <td contenteditable="true"
            oninput="updateCell(${rIndex}, ${cIndex}, this.innerText)">
            ${cell}
        </td>`;
    });
    tr += "</tr>";
    tbody.innerHTML += tr;
  });
}

function updateCell(row, col, value) {
  data[row][col] = value;
}

function addRow() {
  let newRow = new Array(columns.length).fill("");
  newRow[0] = data.length + 1; // auto number
  data.push(newRow);
  renderTable();
}

function addColumn() {
  let name = prompt("Nama Kolom:");
  if (!name) return;

  columns.push(name);
  data.forEach(row => row.push(""));
  renderTable();
}

renderTable();