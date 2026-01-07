let data1 = [];
let data2 = [];
let headers = [];

document.getElementById("file1").addEventListener("change", e => loadFile(e, 1));
document.getElementById("file2").addEventListener("change", e => loadFile(e, 2));

function loadFile(event, fileNo) {
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (fileNo === 1) data1 = json;
    else data2 = json;

    headers = Object.keys(json[0]);
    populateDropdowns(headers);
  };
  reader.readAsBinaryString(event.target.files[0]);
}

function populateDropdowns(headers) {
  ["keyColumn", "qty1", "qty2"].forEach(id => {
    const select = document.getElementById(id);
    select.innerHTML = "";
    headers.forEach(h => {
      const opt = document.createElement("option");
      opt.value = h;
      opt.text = h;
      select.appendChild(opt);
    });
  });
}

function normalize(val) {
  return val.toString().trim().toUpperCase();
}

function runReconciliation() {
  const key = document.getElementById("keyColumn").value;
  const q1 = document.getElementById("qty1").value;
  const q2 = document.getElementById("qty2").value;
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
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Reconciliation");
  XLSX.writeFile(wb, "Reconciliation_Result.xlsx");
}
