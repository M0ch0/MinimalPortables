class XLSXViewer {
    constructor() {
        this.fileInput = document.getElementById('fileInput');
        this.sheetContainer = document.getElementById('sheetContainer');
        this.fontSizeSlider = document.getElementById('fontSizeSlider');
        this.resetButton = document.getElementById('resetButton');
        this.toggleThemeButton = document.getElementById('toggleThemeButton');
        this.copyWithHeaderCheckbox = document.getElementById('copyWithHeader');
        this.enableRangeSelectionCheckbox = document.getElementById('enableRangeSelection');
        this.sheetSelector = document.getElementById('sheetSelector');

        this.isSelectingMultiple = false;
        this.initialCell = null;
        this.workbook = null;

        this.fileInput.addEventListener('change', this.handleFile.bind(this));
        this.sheetSelector.addEventListener('change', this.changeSheet.bind(this));
        this.fontSizeSlider.addEventListener('input', this.updateFontSize.bind(this));
        this.resetButton.addEventListener('click', this.resetView.bind(this));
        this.toggleThemeButton.addEventListener('click', this.toggleTheme.bind(this));
        this.sheetContainer.addEventListener('mousedown', this.handleMouseDown.bind(this));
        this.sheetContainer.addEventListener('dblclick', this.clearSelection.bind(this));
        document.addEventListener('keydown', this.handleCopy.bind(this));

        this.loadSettings();
    }

    handleFile(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => this.processFile(e.target.result);
            reader.readAsArrayBuffer(file);
        }
    }

    processFile(data) {
        this.workbook = XLSX.read(data, { type: 'array' });
        const sheetNames = this.workbook.SheetNames;

        this.sheetSelector.innerHTML = '';

        sheetNames.forEach((name, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.text = name;
            this.sheetSelector.appendChild(option);
        });

        this.displaySheet(this.workbook.Sheets[sheetNames[0]]);
    }

    changeSheet() {
        if (this.workbook) {
            const selectedSheet = this.workbook.Sheets[this.workbook.SheetNames[this.sheetSelector.value]];
            this.displaySheet(selectedSheet);
        }
    }

    displaySheet(sheet) {
        const html = XLSX.utils.sheet_to_html(sheet, { id: "dataTable" });
        this.sheetContainer.innerHTML = html;
        this.applyFontSize();
    }

    updateFontSize() {
        const fontSize = this.fontSizeSlider.value;
        localStorage.setItem('fontSize', fontSize);
        this.applyFontSize();
    }

    applyFontSize() {
        const fontSize = localStorage.getItem('fontSize') || '16';
        const table = document.getElementById('dataTable');
        if (table) {
            table.style.fontSize = `${fontSize}px`;
        }
    }

    resetView() {
        this.sheetContainer.innerHTML = '';
        this.fileInput.value = null;
        this.sheetSelector.innerHTML = '';
        this.workbook = null;
    }


    toggleTheme() {
        document.body.classList.toggle('dark-mode');
        const isDarkMode = document.body.classList.contains('dark-mode');
        localStorage.setItem('darkMode', isDarkMode);
    }

    loadSettings() {
        const savedFontSize = localStorage.getItem('fontSize');
        if (savedFontSize) {
            this.fontSizeSlider.value = savedFontSize;
            this.applyFontSize();
        }

        const darkModeEnabled = JSON.parse(localStorage.getItem('darkMode'));
        if (darkModeEnabled) {
            document.body.classList.add('dark-mode');
        }
    }

    handleMouseDown(e) {
        if (this.enableRangeSelectionCheckbox.checked) {
            if (e.target.tagName === 'TD' || e.target.tagName === 'TH') {
                this.initialCell = e.target;
                this.sheetContainer.style.userSelect = 'text';
    
                document.onmousemove = (e) => {
                    if (e.target.tagName === 'TD' || e.target.tagName === 'TH') {
                        if (this.initialCell !== e.target) {
                            this.isSelectingMultiple = true;
                            this.highlightCells(this.initialCell, e.target);
                            this.sheetContainer.style.userSelect = 'none';
                        }
                    }
                };
    
                document.onmouseup = () => {
                    document.onmousemove = null;
                    document.onmouseup = null;
    
                    if (!this.isSelectingMultiple) {
                        this.clearHighlighting();
                        this.sheetContainer.style.userSelect = 'text';
                    }
    
                    this.isSelectingMultiple = false;
                };
            }
        } else {
            this.sheetContainer.style.userSelect = 'text';
        }
    }
    

    highlightCells(start, end) {
        let startRow = Math.min(start.parentElement.rowIndex, end.parentElement.rowIndex);
        let endRow = Math.max(start.parentElement.rowIndex, end.parentElement.rowIndex);
        let startCol = Math.min(start.cellIndex, end.cellIndex);
        let endCol = Math.max(start.cellIndex, end.cellIndex);

        let table = start.closest('table');
        for (let i = 0; i < table.rows.length; i++) {
            for (let j = 0; j < table.rows[i].cells.length; j++) {
                let cell = table.rows[i].cells[j];
                cell.classList.remove('highlight');
                if (i >= startRow && i <= endRow && j >= startCol && j <= endCol) {
                    cell.classList.add('highlight');
                }
            }
        }
    }

    clearHighlighting() {
        const highlightedCells = document.querySelectorAll('.highlight');
        highlightedCells.forEach(cell => cell.classList.remove('highlight'));
    }

    clearSelection() {
        const selectedCells = document.querySelectorAll('.highlight');
        selectedCells.forEach(cell => cell.classList.remove('highlight'));
    }

    handleCopy(e) {
        if (e.ctrlKey && e.key === 'c') {
            const selectedCells = document.querySelectorAll('.highlight');
            if (selectedCells.length === 0) return;

            let copiedData = '';
            const copyWithHeader = this.copyWithHeaderCheckbox.checked;
            const table = selectedCells[0].closest('table');
            const headerRow = table.rows[0];

            let minCol = Infinity;
            let maxCol = -Infinity;
            selectedCells.forEach(cell => {
                minCol = Math.min(minCol, cell.cellIndex);
                maxCol = Math.max(maxCol, cell.cellIndex);
            });

            if (copyWithHeader) {
                for (let i = minCol; i <= maxCol; i++) {
                    copiedData += headerRow.cells[i].innerText + '\t';
                }
                copiedData = copiedData.trim() + '\n';
            }

            let lastRow = selectedCells[0].parentElement;
            selectedCells.forEach(cell => {
                if (cell.parentElement !== lastRow) {
                    copiedData = copiedData.trim() + '\n';
                    lastRow = cell.parentElement;
                }
                copiedData += cell.innerText + '\t';
            });

            navigator.clipboard.writeText(copiedData.trim());
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new XLSXViewer();
});

document.getElementById('fileInput').addEventListener('change', function() {
    const fileName = this.files[0].name;
    const label = document.querySelector('.file-upload-label');
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
