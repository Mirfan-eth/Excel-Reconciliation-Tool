let data1 = [];
let data2 = [];
let headers1 = [];
let headers2 = [];

document.getElementById("file1").addEventListener("change", e => loadFile(e, 1));
document.getElementById("file2").addEventListener("change", e => loadFile(e, 2));

function loadFile(event, fileNo) {
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (fileNo === 1) {
      data1 = json;
      if (json.length > 0) headers1 = Object.keys(json[0]);
    } else {
      data2 = json;
      if (json.length > 0) headers2 = Object.keys(json[0]);
    }

    // Re-populate dropdowns only when BOTH files are loaded
    if (data1.length > 0 && data2.length > 0) {
      populateDropdowns();
    }
  };
  reader.readAsBinaryString(event.target.files[0]);
}

function populateDropdowns() {
  // Key Column: Show common columns first, then all from both files
  const allKeys = [...new Set([...headers1, ...headers2])];
  
  // Quantity columns: Show columns from respective files
  populateSelect("keyColumn", allKeys);
  populateSelect("qty1", headers1);
  populateSelect("qty2", headers2);
}

function populateSelect(selectId, headerList) {
  const select = document.getElementById(selectId);
  select.innerHTML = "<option value=''>Select...</option>"; // Optional: default option
  
  headerList.forEach(h => {
    const opt = document.createElement("option");
    opt.value = h;
    opt.text = h;
    select.appendChild(opt);
  });
}

function normalize(val) {
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
    alert("Please select all required columns!");
    return;
  }

  const ignoreZero = document.getElementById("ignoreZero").checked;

  let map1 = {};
  let map2 = {};

  data1.forEach(r => {
    const k = normalize(r[key]);
    map1[k] = (map1[k] || 0) + Number(r[q1] || 0);
  });

  data2.forEach(r => {
    const k = normalize(r[key]);
    map2[k] = (map2[k] || 0) + Number(r[q2] || 0);
  });

  let result = [];
  const keys = new Set([...Object.keys(map1), ...Object.keys(map2)]);

  keys.forEach(k => {
    const v1 = map1[k] || 0;
    const v2 = map2[k] || 0;
    const variance = v1 - v2;

    if (ignoreZero && variance === 0) return;

    let status = "Match";
    if (!map2[k]) status = "Missing in File 2";
    else if (!map1[k]) status = "Excess in File 2";
    else if (variance < 0) status = "Short";
    else if (variance > 0) status = "Excess";

    result.push({
      Key: k,
      File1_Qty: v1,
      File2_Qty: v2,
      Variance: variance,
      Status: status
    });
  });

  downloadExcel(result);
}

function downloadExcel(data) {
  if (data.length === 0) {
    alert("No variances found or all variances ignored.");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Reconciliation");
  XLSX.writeFile(wb, "Reconciliation_Result.xlsx");
}

// Credit
console.log("%cExcel Reconciliation Tool - Built by Mohd Irfan", "color: #667eea; font-size: 14px; font-weight: bold;");
