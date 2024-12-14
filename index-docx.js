const fs = require('fs');
const XLSX = require('xlsx');
const { Document, Packer, Paragraph, TextRun } = require('docx');

// Read the Excel file
const workbook = XLSX.readFile('input.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Create a new Word document
const doc = new Document({
  sections: [],
});

data.forEach((row, index) => {
  if (index === 0) return; // Skip header row

  const [level1, level2, level3, content1, content2] = row;

  if (level1) {
    doc.addSection({
      children: [
        new Paragraph({
          children: [new TextRun({ text: level1, bold: true, size: 48 })],
        }),
      ],
    });
  }

  if (level2) {
    doc.addSection({
      children: [
        new Paragraph({
          children: [new TextRun({ text: level2, bold: true, size: 40 })],
        }),
      ],
    });
  }

  if (level3) {
    doc.addSection({
      children: [
        new Paragraph({
          children: [new TextRun({ text: level3, bold: true, size: 32 })],
        }),
      ],
    });
  }

  if (content1 || content2) {
    doc.addSection({
      children: [
        new Paragraph({
          children: [new TextRun(`${content1 || ''} ${content2 || ''}`)],
        }),
      ],
    });
  }
});

// Save the document
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('output.docx', buffer);
  console.log('Word document created successfully.');
}).catch((error) => {
  console.error('Error creating Word document:', error);
});
