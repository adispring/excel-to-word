const fs = require('fs');
const XLSX = require('xlsx');
const officegen = require('officegen');

// Read the Excel file
const workbook = XLSX.readFile('input.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Create a new Word document
const doc = officegen('docx');

data.forEach((row, index) => {
  if (index === 0) return; // Skip header row

  const [level1, level2, level3, content1, content2] = row;

  if (level1) {
    const pObj = doc.createP();
    pObj.addText(level1, { bold: true, font_size: 24 });
  }

  if (level2) {
    const pObj = doc.createP();
    pObj.addText(level2, { bold: true, font_size: 20 });
  }

  if (level3) {
    const pObj = doc.createP();
    pObj.addText(level3, { bold: true, font_size: 16 });
  }

  if (content1 || content2) {
    const pObj = doc.createP();
    pObj.addText(`${content1 || ''} ${content2 || ''}`);
  }
});

// Save the document
const out = fs.createWriteStream('output.docx');
doc.generate(out);
out.on('close', () => {
  console.log('Word document created successfully.');
});
