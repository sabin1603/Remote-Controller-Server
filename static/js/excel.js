let currentWorkbook = null;

function openExcel() {
    hideAllModules();
    fetch('/api/excel/controls/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('excel').style.display = 'flex';
                document.getElementById('excel-controls').style.display = 'flex';
            } else {
                showAlert('Error loading Excel controls.');
            }
        })
        .catch(handleApiError);
}

function openWorkbook(filePath) {
    fetch(`/api/excel/open_workbook/${encodeURIComponent(filePath)}/`, { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message === "Workbook loaded successfully.") {
                currentWorkbook = filePath;
                openExcel();
                showAlert('Workbook opened.');
            } else {
                showAlert('Error opening workbook: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function bringExcelToFront(){
    if (!currentWorkbook){
        showAlert("No workbook open.");
        return;
    }
    fetch('/api/excel/bring_to_front/')
    .then(handleApiResponse)
    .then(data => {
        if (data.status === 'success') {
            console.log('Excel brought to front successfully.');
        } else {
            console.error('Failed to bring Excel to front.');
        }
    })
    .catch(handleApiError);

}

function closeWorkbook() {
    if (!currentWorkbook){
        showAlert("No workbook open.");
        return;
    }
    fetch('/api/excel/close/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message === "Workbook closed successfully.") {
                currentWorkbook = null;
                showAlert('Workbook closed.');
            } else {
                showAlert('Error closing workbook: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function nextWorksheet() {
    fetch('/api/excel/next_worksheet/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function previousWorksheet() {
    fetch('/api/excel/previous_worksheet/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function zoomIn() {
    fetch('/api/excel/zoom_in/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function zoomOut() {
    fetch('/api/excel/zoom_out/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function scrollUp() {
    fetch('/api/excel/scroll_up/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function scrollDown() {
    fetch('/api/excel/scroll_down/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function customScrollLeft() {
    fetch('/api/excel/scroll_left/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}

function scrollRight() {
    fetch('/api/excel/scroll_right/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            console.log(data.message);
        })
        .catch(handleApiError);
}