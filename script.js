
function excelDateToJSDate(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
}

function generaReport() {
  const fileInput = document.getElementById("fileInput");
  const campaignType = document.getElementById("campaignSelect").value;
  const startDateStr = document.getElementById("startDate").value;
  const endDateStr = document.getElementById("endDate").value;
  const output = document.getElementById("output");
  output.textContent = "";

  if (!fileInput.files.length) {
    alert("Carica un file Excel.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets["Raccolta dati"];
    const json = XLSX.utils.sheet_to_json(sheet);

    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    if (isNaN(startDate) || isNaN(endDate)) {
      output.textContent = "⚠️ ATTENZIONE: date non valide. Usa il calendario.";
      return;
    }

    const filtered = json.filter(row => {
      let rawDate = row["Data"];
      if (!rawDate) return false;

      if (typeof rawDate === "number") {
        rawDate = excelDateToJSDate(rawDate);
      } else if (typeof rawDate === "string") {
        const [yyyy, mm, dd] = rawDate.split("T")[0].split("-");
        rawDate = new Date(yyyy, mm - 1, dd);
      } else {
        rawDate = new Date(rawDate);
      }

      if (isNaN(rawDate) || rawDate < startDate || rawDate > endDate) return false;

      const campagna = (row["Campagna"] || "").toLowerCase().trim();
      if (campaignType === "editoriale" && campagna !== "editoriale") return false;
      if (campaignType === "campagna" && campagna !== "campagna") return false;
      if (campaignType === "adv" && !["editoriale", "campagna"].includes(campagna)) return false;

      return true;
    }).map(r => {
      const canale = (r["Canale"] || "").toLowerCase();
      if (["facebook postemobile", "facebook postepay", "instagram postemobile", "instagram postepay"].includes(canale)) {
        r["Canale"] = "Facebook";
      }
      return r;
    });

    window.filteredData = filtered;

    output.textContent = filtered.length ? JSON.stringify(filtered, null, 2) : "Nessun dato trovato.";
  };

  reader.readAsArrayBuffer(fileInput.files[0]);
}

function esportaCSV() {
  if (!window.filteredData || window.filteredData.length === 0) {
    alert("Nessun dato da esportare.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(window.filteredData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, "report_filtrato.csv");
}

function esportaCSVFormattato() {
  if (!window.filteredData || window.filteredData.length === 0) {
    alert("Nessun dato da esportare.");
    return;
  }

  const grouped = {};

  window.filteredData.forEach(row => {
    const label = (row["Label"] || "SENZA LABEL").toUpperCase();
    const argomento = row["Argomento"] || "(Senza argomento)";
    const canale = (row["Canale"] || "").toLowerCase();
    const istituzionale = row["Istituzionale"] ? "✓" : "";

    if (!grouped[label]) grouped[label] = {};
    if (!grouped[label][argomento]) {
      grouped[label][argomento] = {
        "Argomento": argomento,
        "Flag Istituzionale": istituzionale,
        "Facebook": 0,
        "Instagram": 0,
        "LinkedIn": 0,
        "Twitter": 0,
        "YouTube": 0,
        "Totale": 0
      };
    }

    const rowObj = grouped[label][argomento];

    if (canale.includes("facebook")) rowObj["Facebook"] += 1;
    else if (canale.includes("instagram")) rowObj["Instagram"] += 1;
    else if (canale.includes("linkedin")) rowObj["LinkedIn"] += 1;
    else if (canale.includes("twitter") || canale === "x") rowObj["Twitter"] += 1;
    else if (canale.includes("youtube")) rowObj["YouTube"] += 1;

    rowObj["Totale"] =
      rowObj["Facebook"] +
      rowObj["Instagram"] +
      rowObj["LinkedIn"] +
      rowObj["Twitter"] +
      rowObj["YouTube"];
  });

  const finalData = [];
  for (const label in grouped) {
    const somma = {
      "Argomento": label,
      "Flag Istituzionale": "",
      "Facebook": 0,
      "Instagram": 0,
      "LinkedIn": 0,
      "Twitter": 0,
      "YouTube": 0,
      "Totale": 0
    };

    const argomenti = grouped[label];
    for (const argomento in argomenti) {
      const row = argomenti[argomento];

      for (const canale of ["Facebook", "Instagram", "LinkedIn", "Twitter", "YouTube"]) {
        if (row[canale] === 0) row[canale] = "";
        else somma[canale] += typeof row[canale] === "number" ? row[canale] : 0;
      }

      row["Totale"] = ["Facebook", "Instagram", "LinkedIn", "Twitter", "YouTube"]
        .map(k => (typeof row[k] === "number" ? row[k] : 0))
        .reduce((a, b) => a + b, 0);
      somma["Totale"] += row["Totale"];
    }

    finalData.push(somma);
    for (const argomento in argomenti) {
      finalData.push(argomenti[argomento]);
    }
  }

  const ws = XLSX.utils.json_to_sheet(finalData, {
    header: ["Argomento", "Flag Istituzionale", "Facebook", "Instagram", "LinkedIn", "Twitter", "YouTube", "Totale"]
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, "report_editoriale_formattato.csv");
}
