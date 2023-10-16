const fs = require('fs').promises;
const htmlToDocx = require('html-to-docx');
const path = require('path'); // Import the 'path' module

async function convertToDocx(htmlFilePath) {
  try {
    // Read the uploaded HTML file
    const htmlContent = await fs.readFile(htmlFilePath, 'utf-8');

    // Convert HTML to DOCX
    const docxBuffer = await htmlToDocx(htmlContent);

    // Specify the path and filename for the DOCX file
    const docxFilePath = path.join(__dirname, 'output.docx');

    // Save the DOCX file to the server
    await fs.writeFile(docxFilePath, docxBuffer);
    console.log(`DOCX file saved to ${docxFilePath}`);

    return docxBuffer; // Return the DOCX buffer for sending as a response
  } catch (err) {
    console.error('Error:', err);
    throw err;
  }
}

// Export the convertToDocx function to be used in server.js
module.exports = convertToDocx;
