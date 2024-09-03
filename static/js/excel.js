let currentWorkbook = null;

// Function to open Excel and set up the controls
function openExcel() {
    hideAllModules();
    fetch('/api/excel/controls/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('excel').style.display = 'flex';
                document.getElementById('excel-controls').style.display = 'flex';

                // Set the current workbook if available
                if (data.currentWorkbook) {
                    currentWorkbook = data.currentWorkbook; // Assuming server returns current workbook info
                }
            } else {
                alert('Error loading Excel controls.');
            }
        })
        .catch(error => {
            console.error('Error loading Excel controls:', error);
            alert('Error loading Excel controls.');
        });
}

// Function to open a specific workbook
function openWorkbook(fileName) {
    fetch(`/api/excel/open_workbook/?file=${encodeURIComponent(fileName)}`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                currentWorkbook = fileName;  // Set the current workbook
                console.log(`Opened workbook: ${currentWorkbook}`);
                openExcel();  // Load controls once workbook is opened
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening workbook:', error);
            alert('Error opening workbook');
        });
}

// Update this to check if `currentWorkbook` is set before performing any action
function nextWorksheet() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/next_worksheet/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Moved to next worksheet')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error moving to the next worksheet:', error);
            alert('Error moving to the next worksheet');
        });
}

function previousWorksheet() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/previous_worksheet/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Moved to previous worksheet')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error moving to the previous worksheet:', error);
            alert('Error moving to the previous worksheet');
        });
}

function zoomIn() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/zoom_in/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
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

function zoomOut() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/zoom_out/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
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

function scrollUp() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/scroll_up/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
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

function scrollDown() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/scroll_down/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
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

function customScrollLeft() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/scroll_left/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Scrolled left')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error scrolling left:', error);
            alert('Error scrolling left');
        });
}

function scrollRight() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch(`/api/excel/scroll_right/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Scrolled right')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error scrolling right:', error);
            alert('Error scrolling right');
        });
}

function closeWorkbook() {
    if (!currentWorkbook) {
        alert('No workbook open.');
        return;
    }
    fetch(`/api/excel/close/?workbook=${encodeURIComponent(currentWorkbook)}`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Workbook closed successfully.')) {
                currentWorkbook = null;
                document.querySelectorAll('.file-item').forEach(item => item.classList.remove('selected'));
                document.getElementById('workbook-list-container').style.display = 'block';
                document.getElementById('excel-controls').style.display = 'none';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error closing workbook:', error);
            alert('Error closing workbook');
        });
}