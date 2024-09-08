let currentWorkbook = null;

function openExcel() {
    hideAllModules();
    fetch('/api/excel/controls/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('excel').style.display = 'flex';
                document.getElementById('excel-controls').style.display = 'flex';  // Ensure controls are visible
            } else {
                alert('Error loading Excel controls.');
            }
        })
        .catch(error => {
            console.error('Error loading Excel controls:', error);
            alert('Error loading Excel controls.');
        });
}

function openWorkbook(filePath) {
    fetch(`/api/excel/open_workbook/${encodeURIComponent(filePath)}/`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message === "Workbook loaded successfully.") {
                currentWorkbook = filePath;
                openExcel();
                alert('Workbook opened.');
            } else {
                alert('Error opening workbook: ' + data.message);
            }
        })
        .catch(error => {
            console.error('Error opening workbook:', error);
            alert('Error opening workbook.');
        });
}

function closeWorkbook() {
    fetch('/api/excel/close/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message === "Workbook closed successfully.") {
                currentWorkbook = null;
                alert('Workbook closed.');
            } else {
                alert('Error closing workbook: ' + data.message);
            }
        })
        .catch(error => {
            console.error('Error closing workbook:', error);
            alert('Error closing workbook.');
        });
}

function nextWorksheet() {
    fetch('/api/excel/next_worksheet/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error moving to next worksheet:', error);
            alert('Error moving to next worksheet.');
        });
}

function previousWorksheet() {
    fetch('/api/excel/previous_worksheet/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error moving to previous worksheet:', error);
            alert('Error moving to previous worksheet.');
        });
}

function zoomIn() {
    fetch('/api/excel/zoom_in/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error zooming in:', error);
            alert('Error zooming in.');
        });
}

function zoomOut() {
    fetch('/api/excel/zoom_out/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error zooming out:', error);
            alert('Error zooming out.');
        });
}

function scrollUp() {
    fetch('/api/excel/scroll_up/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error scrolling up:', error);
            alert('Error scrolling up.');
        });
}

function scrollDown() {
    fetch('/api/excel/scroll_down/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error scrolling down:', error);
            alert('Error scrolling down.');
        });
}

function customScrollLeft() {
    fetch('/api/excel/scroll_left/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error scrolling left:', error);
            alert('Error scrolling left.');
        });
}

function scrollRight() {
    fetch('/api/excel/scroll_right/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            alert(data.message);
        })
        .catch(error => {
            console.error('Error scrolling right:', error);
            alert('Error scrolling right.');
        });
}