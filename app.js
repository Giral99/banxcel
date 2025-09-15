// ===== Sélecteurs
const pdfFileInput = document.getElementById('pdfFile');
const exportBtn = document.getElementById('exportExcel');
const previewDiv = document.getElementById('preview');
const tolInput = document.getElementById('tol');

let tableRows = []; // sera un tableau 2D (rows x cols)

// Utilitaire : regroupe des objets (x,y,text) par ligne en utilisant une tolérance sur y
function groupByRows(items, yTol = 3) {
  // On trie par y décroissant (dans pdf.js, l'origine peut varier selon les PDFs)
  items.sort((a, b) => b.y - a.y);

  const rows = [];
  let current = [];

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (current.length === 0) {
      current.push(it);
      continue;
    }
    const yRef = current[0].y;
    if (Math.abs(it.y - yRef) <= yTol) {
      current.push(it);
    } else {
      // nouvelle ligne
      rows.push(current);
      current = [it];
    }
  }
  if (current.length) rows.push(current);

  // Trie chaque ligne par x croissant
  rows.forEach(line => line.sort((a, b) => a.x - b.x));
  return rows;
}

// Utilitaire : à partir de la première ligne (entêtes), crée les "centres de colonnes"
function inferColumnCentersFromHeader(headerLine) {
  // On prend le x de chaque fragment comme un centre initial
  const centers = headerLine.map(cell => cell.x);

  // On "nettoie" : si des x sont très proches (< 15 px), on les fusionne (moyenne)
  const merged = [];
  const threshold = 15; // px
  centers.sort((a, b) => a - b);

  let bucket = [centers[0]];
  for (let i = 1; i < centers.length; i++) {
    const prev = bucket[bucket.length - 1];
    const cur = centers[i];
    if (Math.abs(cur - prev) <= threshold) {
      bucket.push(cur);
    } else {
      // push moyenne du bucket
      merged.push(Math.round(bucket.reduce((s, v) => s + v, 0) / bucket.length));
      bucket = [cur];
    }
  }
  if (bucket.length) merged.push(Math.round(bucket.reduce((s, v) => s + v, 0) / bucket.length));

  return merged;
}

// Place chaque fragment sur la colonne dont le centre est le plus proche
function mapLineToColumns(lineItems, centers) {
  const cols = new Array(centers.length).fill('');
  lineItems.forEach(({ x, text }) => {
    let bestIdx = 0;
    let bestDist = Math.abs(x - centers[0]);
    for (let i = 1; i < centers.length; i++) {
      const d = Math.abs(x - centers[i]);
      if (d < bestDist) {
        bestDist = d;
        bestIdx = i;
      }
    }
    cols[bestIdx] = (cols[bestIdx] ? (cols[bestIdx] + ' ' + text) : text).trim();
  });
  return cols;
}

// Lecture d'une page -> renvoie tous les fragments (x,y,text)
async function extractItemsFromPage(page) {
  const textContent = await page.getTextContent();
  // Chaque item possède i.str (texte) et i.transform (matrice). x = transform[4], y = transform[5]
  return textContent.items
    .filter(i => i.str && i.str.trim().length > 0)
    .map(i => {
      const [a, b, c, d, e, f] = i.transform; // e=x, f=y
      return { text: i.str.trim(), x: e, y: f };
    });
}

// Pipeline principal : PDF -> tableRows
async function processPDF(file) {
  const arrayBuf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuf }).promise;

  let allItems = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const items = await extractItemsFromPage(page);
    allItems = allItems.concat(items);
  }

  // 1) Groupement en lignes
  const yTol = parseInt(tolInput.value || '3', 10);
  const rawLines = groupByRows(allItems, yTol).filter(line => line.length > 0);

  // 2) Déterminer la première "vraie" ligne d'entêtes (la plus longue en nb d'items)
  const headerLine = rawLines.reduce((best, cur) => (cur.length > (best?.length || 0) ? cur : best), null);
  if (!headerLine || headerLine.length < 2) {
    throw new Error("Impossible d'inférer les colonnes (pas assez d'éléments sur la ligne d'en-tête). Augmente la tolérance ou essaie un autre PDF tabulaire.");
  }

  // 3) Centres de colonnes
  const centers = inferColumnCentersFromHeader(headerLine);

  // 4) Construire tableRows
  const rows = rawLines.map(line => mapLineToColumns(line, centers));

  // 5) Nettoyage simple : enlever lignes vides et trim
  const cleaned = rows
    .map(r => r.map(c => c.replace(/\s+/g, ' ').trim()))
    .filter(r => r.some(c => c && c.length > 0));

  return cleaned;
}

// Affiche un aperçu (10 lignes)
function renderPreview(rows) {
  const head = rows[0] || [];
  const htmlHead = '<tr>' + head.map(h => `<th>${h || ''}</th>`).join('') + '</tr>';
  const htmlBody = rows.slice(1, 11).map(r => '<tr>' + r.map(c => `<td>${c || ''}</td>`).join('') + '</tr>').join('');
  previewDiv.innerHTML = `<h3>Prévisualisation</h3><div style="overflow:auto;max-height:260px"><table border="1">${htmlHead}${htmlBody}</table></div>`;
}

// Export en Excel
function exportToExcel(rows) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Feuille1");
  XLSX.writeFile(wb, "releve_converti.xlsx");
}

// === Listeners
pdfFileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  previewDiv.innerHTML = "<p>Lecture et reconstruction…</p>";
  exportBtn.disabled = true;

  try {
    tableRows = await processPDF(file);
    renderPreview(tableRows);
    exportBtn.disabled = false;
  } catch (err) {
    previewDiv.innerHTML = `<p style="color:#b00020">Erreur : ${err.message}</p>`;
    exportBtn.disabled = true;
  }
});

exportBtn.addEventListener('click', () => {
  if (tableRows && tableRows.length) exportToExcel(tableRows);
});
