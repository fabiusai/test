
function generaReport() {
  const fileInput = document.getElementById("fileInput");
  const startDate = new Date(document.getElementById("startDate").value);
  const endDate = new Date(document.getElementById("endDate").value);
  const campaignType = document.getElementById("campaignSelect").value;

  if (!fileInput.files.length || isNaN(startDate) || isNaN(endDate)) {
    alert("Carica un file Excel e seleziona entrambe le date.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets["Raccolta dati"];
    const json = XLSX.utils.sheet_to_json(sheet);

    const parseItalianDate = (value) => {
      if (Object.prototype.toString.call(value) === "[object Date]") return value;
      if (typeof value === "string" && value.includes("/")) {
        const [gg, mm, aaaa] = value.split("/");
        return new Date(`${aaaa}-${mm}-${gg}`);
      }
      return new Date(value);
    };

    const filtered = json.filter(row => {
      const dataRow = parseItalianDate(row["Data"]);
      const campagna = (row["Campagna"] || "").toString().toLowerCase().trim();
      const dateOk = dataRow >= startDate && dataRow <= endDate;
      if (!dateOk) return false;

      if (campaignType === "editoriale") return campagna === "editoriale";
      if (campaignType === "campagna") return campagna === "campagna";
      return true;
    });

    console.log("Righe totali:", json.length);
    console.log("Righe filtrate:", filtered.length);

    if (filtered.length === 0) {
      alert("Nessun dato trovato nel periodo selezionato.");
      return;
    }

    filtered.forEach(r => {
      if (r["Canale"]) r["Canale"] = r["Canale"].toLowerCase().trim();
    });

    const gruppi = {};
    filtered.forEach(row => {
      const label = (row["Label"] || "").toUpperCase();
      const argomento = row["Argomento"] || "";
      const chiave = label + "||" + argomento;
      if (!gruppi[chiave]) {
        gruppi[chiave] = {
          Label: label,
          Argomento: argomento,
          Istituzionale: false,
          facebook: 0, instagram: 0, linkedin: 0, twitter: 0, youtube: 0,
          Totale: 0
        };
      }
      const g = gruppi[chiave];
      if (row["Istituzionale"]) g.Istituzionale = true;
      if (g[row["Canale"]] !== undefined) {
        g[row["Canale"]]++;
        g.Totale++;
      }
    });

    const gruppiArray = Object.values(gruppi);
    gruppiArray.sort((a, b) => a.Label.localeCompare(b.Label) || a.Argomento.localeCompare(b.Argomento));

    const wb = XLSX.utils.book_new();
    const ws_data = [];

    ws_data.push(["Periodo di riferimento: dal " + startDate.toLocaleDateString() + " al " + endDate.toLocaleDateString()]);
    ws_data.push(["Canali considerati: Facebook, Instagram, LinkedIn, Twitter, YouTube"]);
    ws_data.push([]);
    ws_data.push(["Argomento", "Flag Istituzionale", "Facebook", "Instagram", "LinkedIn", "Twitter", "YouTube", "Totale"]);

    let lastLabel = null;
    for (const g of gruppiArray) {
      if (lastLabel !== g.Label) {
        if (lastLabel !== null) ws_data.push([]);
        const somma = campo => gruppiArray.filter(x => x.Label === g.Label).reduce((acc, cur) => acc + cur[campo], 0);
        ws_data.push([
          g.Label,
          "",
          somma("facebook") || "",
          somma("instagram") || "",
          somma("linkedin") || "",
          somma("twitter") || "",
          somma("youtube") || "",
          somma("Totale")
        ]);
        lastLabel = g.Label;
      }

      ws_data.push([
        g.Argomento,
        g.Istituzionale ? "✓" : "",
        g.facebook || "",
        g.instagram || "",
        g.linkedin || "",
        g.twitter || "",
        g.youtube || "",
        g.Totale || ""
      ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    ws["!cols"] = [
      { wch: 76 }, { wch: 14 }, { wch: 7 }, { wch: 7 },
      { wch: 7 }, { wch: 7 }, { wch: 7 }, { wch: 14 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, "Post Editoriali");
    XLSX.writeFile(wb, "report_editoriale.xlsx");
  };

  reader.readAsArrayBuffer(fileInput.files[0]);
}
