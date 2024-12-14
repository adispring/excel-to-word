/* eslint-env node */
const fs = require('fs');
const XLSX = require('xlsx');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  SectionType,
} = require('docx');

// Read the Excel file
const workbook = XLSX.readFile('input.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

// Create a new Word document
const doc = new Document({
  sections: [],
});

// Helper function to create paragraphs with specific text, size, and heading level
const createParagraphs = (text, size, bold = false, heading = null) => {
  return text.split('\n').map(line => new Paragraph({
    children: [
      new TextRun({
        text: line,
        bold,
        size,
      }),
    ],
    heading,
  }));
};

data.forEach((row, index) => {
  if (index === 0) return; // Skip header row

  const [level1, level2, level3, content1, content2, content3] = row;

  console.log('Processing row:', row, index);

  const sectionChildren = [];

  if (level1) {
    sectionChildren.push(
      ...createParagraphs(level1, 48, true, HeadingLevel.HEADING_1)
    );
  }

  if (level2) {
    sectionChildren.push(
      ...createParagraphs(level2, 40, true, HeadingLevel.HEADING_2)
    );
  }

  if (level3) {
    sectionChildren.push(
      ...createParagraphs(level3, 36, true, HeadingLevel.HEADING_3)
    );
  }

  // Concatenate content1 and content3 with "; "
  const concatenatedContent = [content1, content3].filter(Boolean).join('; ');

  // Add the concatenated content to the section
  if (concatenatedContent) {
    sectionChildren.push(...createParagraphs(concatenatedContent, 24));
  }

  // Add the section children to the document
  doc.addSection({
    properties: {
      type: SectionType.CONTINUOUS,
    },
    children: sectionChildren,
  });
});

// Save the document
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('output.docx', buffer);
});
