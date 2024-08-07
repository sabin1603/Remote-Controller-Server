let currentPdf = null;

// Function to open the PDF Reader module
function openPdfReader() {
    hideAllModules();
    loadPdfs();
    document.getElementById('pdf-reader').style.display = 'flex';
}

// Function to load the list of PDF files
function loadPdfs() {
    fetch('/api/pdf-reader/list/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.pdfs) {
                const listContainer = document.getElementById('pdf-list');
                listContainer.innerHTML = '';
                data.pdfs.forEach(pdfFile => {
                    const item = document.createElement('div');
                    item.className = 'file-item';
                    item.innerHTML = pdfFile;
                    item.onclick = () => openPdf(pdfFile);
                    listContainer.appendChild(item);
                });
                document.getElementById('pdf-list-container').style.display = 'block';
            } else {
                alert('Error loading PDFs.');
            }
        })
        .catch(error => {
            console.error('Error loading PDFs:', error);
            alert('Error loading PDFs.');
        });
}

// Function to open a specific PDF file
function openPdf(fileName) {
    fetch(`/api/pdf-reader/open/${encodeURIComponent(fileName)}/`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('PDF loaded successfully')) {
                currentPdf = fileName;
                document.querySelectorAll('.file-item').forEach(item => {
                    item.classList.toggle('selected', item.innerHTML === fileName);
                });
                document.getElementById('pdf-list-container').style.display = 'none';
                document.getElementById('pdf-reader-controls').style.display = 'flex';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening PDF:', error);
            alert('Error opening PDF');
        });
}

// Function to scroll up in the PDF document
function scrollUpPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/scroll_up/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Scrolled up')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error scrolling up:', error);
            alert('Error scrolling up');
        });
}

// Function to scroll down in the PDF document
function scrollDownPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/scroll_down/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Scrolled down')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error scrolling down:', error);
            alert('Error scrolling down');
        });
}

// Function to zoom in on the PDF document
function zoomInPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/zoom_in/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Zoomed in')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error zooming in:', error);
            alert('Error zooming in');
        });
}

// Function to zoom out on the PDF document
function zoomOutPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/zoom_out/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Zoomed out')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error zooming out:', error);
            alert('Error zooming out');
        });
}

// Function to enable read mode in the PDF document
function enableReadModePdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/enable_read_mode/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Read mode enabled')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error enabling read mode:', error);
            alert('Error enabling read mode');
        });
}

// Function to disable read mode in the PDF document
function disableReadModePdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/disable_read_mode/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Read mode disabled')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error disabling read mode:', error);
            alert('Error disabling read mode');
        });
}

// Function to save the PDF document
function savePdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/save/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('PDF saved')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error saving PDF:', error);
            alert('Error saving PDF');
        });
}

// Function to print the PDF document
function printPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf-reader/print/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Print dialog opened')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening print dialog:', error);
            alert('Error opening print dialog');
        });
}

// Function to close the PDF document
function closePdf() {
    if (!currentPdf) {
        alert('No PDF open.');
        return;
    }
    fetch('/api/pdf-reader/close/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('PDF closed')) {
                currentPdf = null;
                document.querySelectorAll('.file-item').forEach(item => item.classList.remove('selected'));
                document.getElementById('pdf-list-container').style.display = 'block';
                document.getElementById('pdf-reader-controls').style.display = 'none';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error closing PDF:', error);
            alert('Error closing PDF');
        });
}
