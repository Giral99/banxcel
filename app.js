// Sélection des éléments
const pdfFileInput = document.getElementById('pdfFile');
const exportBtn = document.getElementById('exportExcel');
const previewDiv = document.getElementById('preview');

let extractedRows = []; // stockera les lignes extraites

// 1. Lecture du PDF dès qu’un fichier est choisi
pdfFileInput.addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  previewDiv.innerHTML = "<p>Lecture du PDF en cours...</p>";
  extractedRows = [];

  const fileReader = new FileReader();
  fileReader.onload = async function() {
    const typedarray = new Uint8Array(this.result);

    // Utilisation de pdf.js pour lire le PDF
    const pdf = await pdfjsLib.getDocument(typedarray).promise;

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();

      // Concatène tout le texte de la page
      let pageText = textContent.items.map(i => i.str).join(' ');
      
      // Découpe grossièrement par lignes (selon ton PDF tu ajusteras)
      let lines = pageText.split(/[\r\n]+/);

      lines.forEach(line => {
        if (line.trim().length > 5) { // ignore lignes trop courtes
          extractedRows.push([line.trim()]);
        }
      });
    }

    if (extractedRows.length > 0) {
      previewDiv.innerHTML = "<h3>Prévisualisation</h3>";
      let table = "<table border='1'><tr><th>Lignes extraites</th></tr>";
      extractedRows.slice(0,10).forEach(row => {
        table += `<tr><td>${row}</td></tr>`;
      });
      table += "</table>";
      previewDiv.innerHTML += table;
      exportBtn.disabled = false;
    } else {
      previewDiv.innerHTML = "<p>Aucune donnée extraite</p>";
      exportBtn.disabled = true;
    }
  };

  fileReader.readAsArrayBuffer(file);
});

// 2. Export Excel avec SheetJS
exportBtn.addEventListener('click', () => {
  const ws = XLSX.utils.aoa_to_sheet(extractedRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Données PDF");
  XLSX.writeFile(wb, "releve_converti.xlsx");
});

