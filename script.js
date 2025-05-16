
function generaReport() {
  const fileInput = document.getElementById("fileInput");
  const campaignType = document.getElementById("campaignSelect").value;

  if (!fileInput.files.length) {
    alert("Carica un file Excel.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    const filtered = json.filter(row => {
      const campagna = (row["Campagna"] || "").toString().toLowerCase().trim();
      if (campaignType === "editoriale") return campagna === "editoriale";
      if (campaignType === "campagna") return campagna === "campagna";
      return true;
    });

    if (filtered.length === 0) {
      alert("Nessun dato disponibile con i criteri selezionati.");
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

    const rows = [];
    rows.push(["Argomento", "Flag Istituzionale", "Facebook", "Instagram", "LinkedIn", "Twitter", "YouTube", "Totale"]);

    let lastLabel = null;
    for (const g of gruppiArray) {
      if (lastLabel !== g.Label) {
        if (lastLabel !== null) rows.push([]);
        const somma = campo => gruppiArray.filter(x => x.Label === g.Label).reduce((acc, cur) => acc + cur[campo], 0);
        rows.push([
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

      rows.push([
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

    const csvContent = "data:text/csv;charset=utf-8," + rows.map(r => r.map(v => `"${v}"`).join(",")).join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "report_editoriale.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  reader.readAsArrayBuffer(fileInput.files[0]);
}
