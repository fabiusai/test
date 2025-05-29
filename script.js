function showTime() {
	document.getElementById('currentTime').innerHTML = new Date().toUTCString();
}
showTime();
setInterval(function () {
	showTime();
}, 1000);
// --- Inizio codice CSV ---

// Funzione per esportare in CSV
function exportCsv() {
  const header = ['tabIndex','wa','ytTitle','ytDesc','ytLink','tgLink'];
  const rows = tabsConfig.map((_, i) => {
    const data = JSON.parse(localStorage.getItem(`posteTab${i}`) || '{}');
    return header.map(h => {
      const v = data[h] || '';
      return `"${v.replace(/"/g, '""')}"`;
    }).join(',');
  });

  const csvContent = [header.join(','), ...rows].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'posteData.csv';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// Funzione per leggere e parsare CSV
function parseCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines[0]
    .split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/)
    .map(h => h.replace(/^"|"$/g, ''));
  return lines.slice(1).map(line => {
    const cols = line
      .split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/)
      .map(v => v.replace(/^"|"$/g, '').replace(/""/g, '"'));
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = cols[idx] || ''; });
    return obj;
  });
}

// Funzione per importare CSV e popolare i campi
function importCsv(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const records = parseCsv(e.target.result);
      records.forEach(rec => {
        const idx = rec.tabIndex;
        localStorage.setItem(`posteTab${idx}`, JSON.stringify(rec));
      });
      initializeApp();
      alert('Importazione CSV completata!');
    } catch (err) {
      console.error(err);
      alert('Errore nel parsing del CSV. Verifica il formato.');
    }
  };
  reader.readAsText(file);
}

// Al caricamento della pagina, colleghiamo i listener
window.addEventListener('DOMContentLoaded', () => {
  document.getElementById('exportCsv').addEventListener('click', exportCsv);

  const importBtn = document.getElementById('importButton');
  const fileInput = document.getElementById('importCsvInput');

  importBtn.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => {
    if (e.target.files.length > 0) {
      importCsv(e.target.files[0]);
      e.target.value = '';
    }
  });
});

// --- Fine codice CSV ---
