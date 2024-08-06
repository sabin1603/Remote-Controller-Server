let currentDocument = null;

// Function to open the Word module
function openWord() {
    hideAllModules();
    loadDocuments();
    document.getElementById('word').style.display = 'flex';
}

// Function to load the list of Word documents
function loadDocuments() {
    fetch('/api/word/list_documents/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.documents) {
                const listContainer = document.getElementById('document-list');
                listContainer.innerHTML = '';
                data.documents.forEach(wordDocument => {
                    const item = document.createElement('div');
                    item.className = 'file-item';
                    item.innerHTML = wordDocument;
                    item.onclick = () => openDocument(wordDocument);
                    listContainer.appendChild(item);
                });
                document.getElementById('document-list-container').style.display = 'block';
            } else {
                alert('Error loading documents.');
            }
        })
        .catch(error => {
            console.error('Error loading documents:', error);
            alert('Error loading documents.');
        });
}

// Function to open a specific Word document
function openDocument(fileName) {
    fetch(`/api/word/open_document/${encodeURIComponent(fileName)}/`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('loaded successfully')) {
                currentDocument = fileName;
                document.querySelectorAll('.file-item').forEach(item => {
                    item.classList.toggle('selected', item.innerHTML === fileName);
                });
                document.getElementById('document-list-container').style.display = 'none';
                document.getElementById('word-controls').style.display = 'flex';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening document:', error);
            alert('Error opening document');
        });
}

// Function to scroll up in the Word document
function scrollUpWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/scroll_up/', { method: 'GET' })
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

// Function to scroll down in the Word document
function scrollDownWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/scroll_down/', { method: 'GET' })
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

// Function to zoom in on the Word document
function zoomInWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/zoom_in/', { method: 'GET' })
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

// Function to zoom out on the Word document
function zoomOutWord() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/zoom_out/', { method: 'GET' })
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

// Function to enable read mode in the Word document
function enableReadMode() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/enable_read_mode/', { method: 'GET' })
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

// Function to disable read mode in the Word document
function disableReadMode() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/disable_read_mode/', { method: 'GET' })
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

// Function to move to the next page in the Word document
function nextPage() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/next_page/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Moved to next page')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error moving to next page:', error);
            alert('Error moving to next page');
        });
}

// Function to move to the previous page in the Word document
function previousPage() {
    if (!currentDocument) {
        alert('No document selected.');
        return;
    }
    fetch('/api/word/previous_page/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Moved to previous page')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error moving to previous page:', error);
            alert('Error moving to previous page');
        });
}

// Function to close the Word document
function closeDocument() {
    if (!currentDocument) {
        alert('No document open.');
        return;
    }
    fetch('/api/word/close_document/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Document closed successfully.')) {
                currentDocument = null;
                document.querySelectorAll('.file-item').forEach(item => item.classList.remove('selected'));
                document.getElementById('document-list-container').style.display = 'block';
                document.getElementById('word-controls').style.display = 'none';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error closing document:', error);
            alert('Error closing document');
        });
}
