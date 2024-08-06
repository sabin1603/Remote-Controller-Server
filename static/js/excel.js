let currentWorkbook = null;

function openExcel() {
    hideAllModules();
    loadWorkbooks();
    document.getElementById('excel').style.display = 'flex';
}

function loadWorkbooks() {
    fetch('/api/excel/list_workbooks/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.workbooks) {
                const listContainer = document.getElementById('workbook-list');
                listContainer.innerHTML = '';
                data.workbooks.forEach(workbook => {
                    const item = document.createElement('div');
                    item.className = 'file-item';
                    item.innerHTML = workbook;
                    item.onclick = () => openWorkbook(workbook);
                    listContainer.appendChild(item);
                });
                document.getElementById('workbook-list-container').style.display = 'block';
            } else {
                alert('Error loading workbooks.');
            }
        })
        .catch(error => {
            console.error('Error loading workbooks:', error);
            alert('Error loading workbooks.');
        });
}

function openWorkbook(fileName) {
    fetch(`/api/excel/open_workbook/${encodeURIComponent(fileName)}/`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('loaded successfully')) {
                currentWorkbook = fileName;
                document.querySelectorAll('.file-item').forEach(item => {
                    item.classList.toggle('selected', item.innerHTML === fileName);
                });
                document.getElementById('workbook-list-container').style.display = 'none';
                document.getElementById('excel-controls').style.display = 'flex';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening workbook:', error);
            alert('Error opening workbook');
        });
}

function nextWorksheet() {
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch('/api/excel/next_worksheet/', { method: 'GET' })
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
    fetch('/api/excel/previous_worksheet/', { method: 'GET' })
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
    fetch('/api/excel/zoom_in/', { method: 'GET' })
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
    fetch('/api/excel/zoom_out/', { method: 'GET' })
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
    fetch('/api/excel/scroll_up/', { method: 'GET' })
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
    fetch('/api/excel/scroll_down/', { method: 'GET' })
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
    console.log('Scroll left button clicked');  // Debug log
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch('/api/excel/scroll_left/', { method: 'GET' })
        .then(response => {
            console.log('Response status:', response.status);  // Debug log
            return response.json();
        })
        .then(data => {
            console.log('Response data:', data);  // Debug log
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
    fetch('/api/excel/scroll_right/', { method: 'GET' })
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

function closeWorkbook(){
    if (!currentWorkbook) {
        alert('No workbook open.');
        return;
    }
    fetch('/api/excel/close/', { method: 'GET' })
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