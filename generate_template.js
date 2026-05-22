// Script untuk generate template Excel nilai
// Run: node generate_template.js

import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Buat template dengan struktur data
const templateData = [
  ['No', 'NIS', 'Nama Santri', 'Matematika - UTS', 'Matematika - UAS', 'B. Indonesia - UTS', 'B. Indonesia - UAS', 'IPA - UTS', 'IPA - UAS'],
  [1, '001', 'Ahmad Abdullah', 85, 88, 75, 80, 90, 92],
  [2, '002', 'Badria Siti', 90, 92, 85, 88, 88, 85],
  [3, '003', 'Cilak Maulana', 78, 82, 70, 75, 82, 80],
  [4, '004', 'Dina Nursita', 88, 85, 80, 85, 85, 88],
  [5, '005', 'Eka Rahman', 92, 95, 90, 93, 95, 97],
  // Baris kosong untuk template
  [6, '006', 'Farah Liana', '', '', '', '', '', ''],
  [7, '007', 'Ghufran Malik', '', '', '', '', '', ''],
];

// Buat workbook
const ws = XLSX.utils.aoa_to_sheet(templateData);

// Set column widths
ws['!cols'] = [
  { wch: 8 },   // No
  { wch: 15 },  // NIS
  { wch: 25 },  // Nama Santri
  { wch: 18 }, // Matematika UTS
  { wch: 18 }, // Matematika UAS
  { wch: 18 }, // B. Indonesia UTS
  { wch: 18 }, // B. Indonesia UAS
  { wch: 12 }, // IPA UTS
  { wch: 12 }, // IPA UAS
];

// Freeze panes (header row + left 3 columns)
ws['!freeze'] = { xSplit: 3, ySplit: 1 };

// Create workbook
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Nilai');

// Save to public folder
const outputPath = path.join(__dirname, 'public', 'template_nilai_sample.xlsx');
XLSX.writeFile(wb, outputPath);

console.log('✓ Template Excel berhasil dibuat di:', outputPath);
console.log('File ini menunjukkan format yang diharapkan untuk impor nilai.');
