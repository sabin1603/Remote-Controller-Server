let currentDocument = null;

function openWord() {
    hideAllModules();
    fetch('/api/word/controls/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('word').style.display = 'flex';
                document.getElementById('word-controls').style.display = 'flex';
            } else {
                showAlert('Error loading Word controls.');
            }
        })
        .catch(handleApiError);
}

function openDocument(filePath) {
    fetch(`/api/word/open_document/${encodeURIComponent(filePath)}/`, { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message === "Document loaded successfully.") {
                currentDocument = filePath;
                openWord();
                showAlert('Document opened.');
            } else {
                showAlert('Error opening document: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function bringWordToFront(){
    fetch('/api/word/bring_to_front/')
    .then(handleApiResponse)
    .then(data => {
        if (data.status === 'success') {
            console.log('Word brought to front successfully.');
        } else {
            console.error('Failed to bring Word to front.');
        }
    })
    .catch(handleApiError);
}

function scrollUpWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/scroll_up/', { method: 'GET' })
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

function scrollDownWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/scroll_down/', { method: 'GET' })
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

function zoomInWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/zoom_in/', { method: 'GET' })
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

function zoomOutWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/zoom_out/', { method: 'GET' })
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

function enableReadMode() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/enable_read_mode/', { method: 'GET' })
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

function disableReadMode() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/disable_read_mode/', { method: 'GET' })
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

function closeDocument() {
    if (!currentDocument) {
        alert('No document open.');
        return;
    }
    fetch('/api/word/close/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Document closed successfully.')) {
                currentDocument = null;
                showAlert(data.message);
            } else {
                alert('Error: ' + data.message);
            }
        })
        .catch(handleApiError);
}
