document.getElementById('sanitizeButton').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    const result = document.getElementById('result');

    if (!fileInput.files.length) {
        result.textContent = 'Please upload an XLSX file first.';
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Sanitize workbook
        sanitizeXLSX(workbook);

        // Create a new sanitized file
        const sanitizedData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([sanitizedData], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);

        // Create a download link
        const downloadLink = document.createElement('a');
        downloadLink.href = url;
        downloadLink.download = 'sanitized_file.xlsx';
        downloadLink.textContent = 'Download Sanitized File';
        downloadLink.style.display = 'block';
        result.innerHTML = '';
        result.appendChild(downloadLink);
    };

    reader.readAsArrayBuffer(file);
});

function sanitizeXLSX(workbook) {
    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];

        for (const cell in sheet) {
            if (cell[0] === '!') continue; // Skip special properties

            // Remove formulas
            if (sheet[cell].f) {
                delete sheet[cell].f; // Remove the formula
                sheet[cell].v = ''; // Optionally clear the value
            }
        }
    }

    // Remove macros, external links, and other potentially harmful content
    removeMacros(workbook);
    removeExternalLinks(workbook);
    removeEmbeddings(workbook);
    removeMetadata(workbook);
}

function removeMacros(workbook) {
    // Remove VBA-related parts
    delete workbook.Workbook?.Names;
    delete workbook.Workbook?.VBAMacros;
}

function removeExternalLinks(workbook) {
    // Remove external connections
    for (const name in workbook.Sheets) {
        const sheet = workbook.Sheets[name];
        for (const key in sheet) {
            if (key.startsWith('!')) continue;
            if (sheet[key].f?.startsWith('HYPERLINK')) {
                delete sheet[key].f;
            }
        }
    }
}

function removeEmbeddings(workbook) {
    // Remove any embedded objects
    const contentTypes = workbook["[Content_Types].xml"];
    if (contentTypes) {
        for (const type of contentTypes) {
            if (type.PartName.includes('embeddings')) {
                delete contentTypes[type.PartName];
            }
        }
    }
}

function removeMetadata(workbook) {
    // Remove metadata from core and app properties
    delete workbook.Props?.Author;
    delete workbook.Props?.LastAuthor;
    delete workbook.Props?.CreatedDate;
    delete workbook.Props?.ModifiedDate;
}


document.getElementById('fileInput').addEventListener('change', function() {
    const fileName = this.files[0].name;
    const label = document.getElementById('fileUploadLabel')
    label.textContent = fileName ? fileName : 'Click or Drag to Upload';
});

const dropArea = document.querySelector('.file-upload'); // Assuming the file upload area has this class

dropArea.addEventListener('dragover', (event) => {
    event.preventDefault(); // Necessary to allow drop
    dropArea.classList.add('highlight'); // Optional visual feedback
});

dropArea.addEventListener('dragleave', () => {
    dropArea.classList.remove('highlight'); // Optional visual feedback
});

dropArea.addEventListener('drop', (event) => {
    event.preventDefault();
    dropArea.classList.remove('highlight'); // Optional visual feedback
    const file = event.dataTransfer.files[0];
    const fileInput = document.getElementById('fileInput');
    fileInput.files = event.dataTransfer.files; // Set the file input's files
    // Trigger the change event to handle the file
    fileInput.dispatchEvent(new Event('change'));
});

