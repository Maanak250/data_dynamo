const express = require('express');
const fileUpload = require('express-fileupload');
const convertToDocx = require('./docxConverter');
const path = require('path');
const fs = require('fs');
const app = express();

const PORT = process.env.PORT || 3000;

const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

app.use(fileUpload());
app.use(express.static(path.join(__dirname, 'uploads')));

app.post('/generate-docx', async (req, res) => {
  if (!req.files || !req.files.htmlFile) {
    return res.status(400).send('No HTML file uploaded.');
  }

  const htmlFile = req.files.htmlFile;
  const htmlFilePath = path.join(__dirname, 'uploads', htmlFile.name);

  try {
    await htmlFile.mv(htmlFilePath);
    const docxFilePath = 'output.docx'; // Define docxFilePath
    await convertToDocx(htmlFilePath);

    // Set response headers for the generated DOCX file and send it as a download
    res.set('Content-Disposition', `attachment; filename=${htmlFile.name}.docx`);
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.download(docxFilePath, (err) => {
      if (err) {
        console.error('Error sending the generated DOCX:', err);
        res.status(500).json({ error: 'Error sending the file' });
      }
    });
  } catch (error) {
    console.error('Error generating DOCX:', error);
    res.status(500).json({ error: 'DOCX generation failed' });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
