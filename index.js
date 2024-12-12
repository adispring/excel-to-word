const fs = require('fs');
const XLSX = require('xlsx');
const { Document, Packer, Paragraph, HeadingLevel } = require('docx');

// Read the Excel file
const workbook = XLSX.readFile('input.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Create a new Word document
const doc = new Document();

let currentLevel1 = null;
let currentLevel2 = null;
let currentLevel3 = null;

data.forEach((row, index) => {
  if (index === 0) return; // Skip header row

  const [level1, level2, level3, content1, content2] = row;

  if (level1) {
    currentLevel1 = new Paragraph({
      text: level1,
      heading: HeadingLevel.HEADING_1,
    });
    doc.addSection({ children: [currentLevel1] });
  }

  if (level2) {
    currentLevel2 = new Paragraph({
      text: level2,
      heading: HeadingLevel.HEADING_2,
    });
    doc.addSection({ children: [currentLevel2] });
  }

  if (level3) {
    currentLevel3 = new Paragraph({
      text: level3,
      heading: HeadingLevel.HEADING_3,
    });
    doc.addSection({ children: [currentLevel3] });
  }

  if (content1 || content2) {
    const content = new Paragraph({
      text: `${content1 || ''} ${content2 || ''}`,
    });
    doc.addSection({ children: [content] });
  }
});

// Save the document
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('output.docx', buffer);
  console.log('Word document created successfully.');
});
