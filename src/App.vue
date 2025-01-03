<template>
  <div class="container">
    <h1>Extracteur PDF vers Excel pour Panier Netstore</h1>

    <div class="upload-section">
      <input type="file" accept=".pdf" @change="handleFileUpload"/>
    </div>

    <div v-if="extractedData.length" class="results-section">
      <h2>Données extraites</h2>
      <table>
        <thead>
        <tr>
          <th>Référence</th>
          <th>Quantité</th>
        </tr>
        </thead>
        <tbody>
        <tr v-for="(item, index) in extractedData" :key="index">
          <td>{{ item.reference }}</td>
          <td>{{ item.quantity }}</td>
        </tr>
        </tbody>
      </table>

      <button @click="exportToExcel" class="export-btn">
        Exporter vers Excel
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import {ref} from 'vue'
import * as XLSX from 'xlsx'
import * as pdfjsLib from 'pdfjs-dist'

pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`

interface ExtractedItem {
  reference: string
  quantity: number
}

const extractedData = ref<ExtractedItem[]>([])

async function handleFileUpload(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (!file) return;

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

  const extractedItems: ExtractedItem[] = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const text = textContent.items.map((item: any) => item.str).join(" ");

    console.log("Contenu brut extrait : ", text);

    const cleanedText = text.replace(/\s+/g, ' ').trim();

    const regex =
        /([A-Z0-9,]{6,40})\s+([A-ZÉÈÊÀ0-9\+\(\)\-\,\²\³\.\/\-\_\s]+(?:\s+[A-ZÉÈÊÀ0-9\\(\)+\-\,\²\³\.\/\-\_\s]+)*)\s+((\d{1,2}\s(?:janv|févr|mars|avr|mai|juin|juil|août|sept|oct|nov|déc)\.)|demain|lundi)\s+(\d+)\s+([A-Z0-9]{1,4})\s+(\d+(?:,\d{2,4})?€?)\s+(\d+(?:,\d{2})?€?)/g

    let match;

    while ((match = regex.exec(cleanedText)) !== null) {
      console.log("Correspondance trouvée : ", match);

      const reference = match[1];
      const quantity = parseInt(match[5], 10);
      if (!isNaN(quantity) && quantity > 0) {
        extractedItems.push({reference, quantity});
      }
    }
  }

  extractedData.value = extractedItems;
}


function exportToExcel() {
  const worksheet = XLSX.utils.json_to_sheet(extractedData.value)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, "Articles")
  XLSX.writeFile(workbook, "articles_extraits.xlsx")
}
</script>

<style>
.container {
  max-width: 800px;
  margin: 0 auto;
  padding: 2rem;
}

.upload-section {
  margin: 2rem 0;
}

.results-section {
  margin-top: 2rem;
}

table {
  width: 100%;
  border-collapse: collapse;
  margin: 1rem 0;
}

th, td {
  border: 1px solid #ddd;
  padding: 0.5rem;
  text-align: left;
}

th {
  background-color: #f5f5f5;
}

.export-btn {
  background-color: #4CAF50;
  color: white;
  padding: 0.5rem 1rem;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.export-btn:hover {
  background-color: #45a049;
}
</style>
