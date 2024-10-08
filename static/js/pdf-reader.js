let currentPdf = null;

function openPdfReader() {
    hideAllModules();
    fetch('/api/pdf_reader/controls/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('pdf-reader').style.display = 'flex';
                document.getElementById('pdf-reader-controls').style.display = 'flex';
            } else {
                showAlert('Error loading PDF Reader controls.');
            }
        })
        .catch(handleApiError);
}

function openPdf(filePath) {
    fetch(`/api/pdf_reader/open/${encodeURIComponent(filePath)}/`, { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message === "PDF loaded successfully.") {
                currentPdf = filePath;
                openPdfReader();
                showAlert('PDF opened.');
            } else {
                showAlert('Error opening PDF: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function bringAcrobatToFront(){
    fetch('/api/pdf_reader/bring_to_front/')
    .then(handleApiResponse)
    .then(data => {
        if (data.status === 'success') {
            console.log('Acrobat brought to front successfully.');
        } else {
            console.error('Failed to bring Acrobat to front.');
        }
    })
    .catch(handleApiError);
}

function scrollUpPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/scroll_up/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Scrolled up')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function scrollDownPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/scroll_down/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Scrolled down')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function zoomInPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/zoom_in/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Zoomed in')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function zoomOutPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/zoom_out/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Zoomed out')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function enableReadModePdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/enable_read_mode/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Read mode enabled')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function disableReadModePdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/disable_read_mode/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Read mode disabled')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function savePdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/save/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('PDF saved')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function printPdf() {
    if (!currentPdf) {
        alert('No PDF selected.');
        return;
    }
    fetch('/api/pdf_reader/print/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Print dialog opened')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(handleApiError);
}

function closePdf() {
    fetch('/api/pdf_reader/close/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes("PDF closed successfully.")) {
                currentPresentation = null;
                showAlert(data.message);
            } else {
                showAlert('Error: ' + data.message);
            }
        })
        .catch(handleApiError);
}
