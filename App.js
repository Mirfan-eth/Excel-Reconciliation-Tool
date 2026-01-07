let data1 = [];
let data2 = [];
let headers1 = [];
let headers2 = [];

document.getElementById("file1").addEventListener("change", e => loadFile(e, 1));
document.getElementById("file2").addEventListener("change", e => loadFile(e, 2));

function loadFile(event, fileNo) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (json.length === 0) {
        alert("Selected Excel file is empty or has no data!");
        return;
      }

      if (fileNo === 1) {
        data1 = json;
        headers1 = Object.keys(json[0]);
      } else {
        data2 = json;
        headers2 = Object.keys(json[0]);
      }

      // Only populate when both files are loaded
      if (data1.length > 0 && data2.length > 0) {
        populateDropdowns();
      }
    } catch (err) {
      alert("Error reading file: " + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

function populateDropdowns() {
  // Key column: show all unique columns from both files
  const allKeys = [...new Set([...headers1, ...headers2])];
  
  populateSelect("keyColumn", allKeys);
  populateSelect("qty1", headers1);
  populateSelect("qty2", headers2);
}

function populateSelect(selectId, headerList) {
  const select = document.getElementById(selectId);
  select.innerHTML = '<option value="">-- Select Column --</option>';
  
  headerList.forEach(header => {
    const option = document.createElement("option");
    option.value = header;
    option.textContent = header;
    select.appendChild(option);
  });
}

function normalize(val) {
  if (val === null || val === undefined) return "";
  return val.toString().trim().toUpperCase();
}

function runReconciliation() {
  if (data1.length === 0 || data2.length === 0) {
    alert("Please upload both Excel files first!");
    return;
  }

  const key = document.getElementById("keyColumn").value;
  const q1 = document.getElementById("qty1").value;
  const q2 = document.getElementById("qty2").value;

  if (!key || !q1 || !q2) {
    alert("Please select Matching Key and Quantity columns for both files!");
    return;
  }

  const ignoreZero = document.getElementById("ignoreZero").checked;

  let map1 = {};
  let map2 = {};

  data1.forEach(row => {
    const k = normalize(row[key]);
    if (k !== "") {
      map1[k] = (map1[k] || 0) + Number(row[q1] || 0);
    }
  });

  data2.forEach(row => {
    const k = normalize(row[key]);
    if (k !== "") {
      map2[k] = (map2[k] || 0) + Number(row[q2] || 0);
    }
  });

  let result = [];
  const allKeys = new Set([...Object.keys(map1), ...Object.keys(map2)]);

  allKeys.forEach(k => {
    const v1 = map1[k] || 0;
    const v2 = map2[k] || 0;
    const variance = v1 - v2;

    if (ignoreZero && variance === 0) return;

    let status = "Match";
    if (v1 === 0) status = "Missing in File 1 (Excess in File 2)";
    else if (v2 === 0) status = "Missing in File 2 (Excess in File 1)";
    else if (variance > 0) status = "Excess in File 1";
    else if (variance < 0) status = "Short in File 1";

    result.push({
      Key: k,
      "File 1 Qty": v1,
      "File 2 Qty": v2,
      Variance: variance,
      Status: status
    });
  });

  if (result.length === 0) {
    alert("No differences found or all variances are zero (ignored).");
    return;
  }

  downloadExcel(result);
}

function downloadExcel(data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Reconciliation");
  XLSX.writeFile(wb, "Reconciliation_Result.xlsx");
}

console.log("%cExcel Reconciliation Tool - Built by Mohd Irfan", "color: #667eea; font-size: 14px; font-weight: bold;");
