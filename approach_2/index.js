const fs = require('fs').promises;
const htmlToDocx = require('html-to-docx');

async function convertAndSave() {
  const htmlContent = `
  
  
  `;

  try {
    const docxBuffer = await htmlToDocx(htmlContent);

    const docxFilePath = 'output.docx';

    await fs.writeFile(docxFilePath, docxBuffer);
    console.log(`DOCX file saved to ${docxFilePath}`);
  } catch (err) {
    console.error('Error:', err);
  }
}

convertAndSave();

