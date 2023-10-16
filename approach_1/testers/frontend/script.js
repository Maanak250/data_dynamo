async function convertToDocx() {
  const htmlFileInput = document.getElementById('htmlFile');
  
  if (!htmlFileInput || !htmlFileInput.files || htmlFileInput.files.length === 0) {
    alert('Please select an HTML file.');
    return;
  }

  const htmlFile = htmlFileInput.files[0];
  const formData = new FormData();
  formData.append('htmlFile', htmlFile);

  try {
    const response = await fetch('http://localhost:3000/generate-docx', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      console.error('Conversion failed');
      return;
    }

    const blob = await response.blob();
    const downloadLink = document.getElementById('downloadLink');
    const downloadButton = document.getElementById('downloadButton');

    downloadButton.href = URL.createObjectURL(blob);
    downloadLink.style.display = 'block';
  } catch (error) {
    console.error('Error:', error);
  }
}
