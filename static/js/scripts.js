let currentPath = '/';
let rootPath = '/';
let history = [];
let historyIndex = -1;
let selectedFolder = null;
let currentPresentation = null;
let currentWorkbook = null;
let selectedFilePath = null;
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

function scrollLeft() {
    console.log('Scroll left button clicked');  // Debug log
    if (!currentWorkbook) {
        alert('No workbook selected.');
        return;
    }
    fetch('/api/excel/scroll_left/', { method: 'GET' })
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


function openPowerpoint() {
    hideAllModules();
    loadPresentations();
    document.getElementById('powerpoint').style.display = 'flex';
}

function loadPresentations() {
    fetch('/api/powerpoint/list_presentations/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.presentations) {
                const listContainer = document.getElementById('presentation-list');
                listContainer.innerHTML = '';
                data.presentations.forEach(presentation => {
                    const item = document.createElement('div');
                    item.className = 'file-item';
                    item.innerHTML = presentation;
                    item.onclick = () => openPresentation(presentation);
                    listContainer.appendChild(item);
                });
                document.getElementById('presentation-list-container').style.display = 'block';
            } else {
                alert('Error loading presentations.');
            }
        })
        .catch(error => {
            console.error('Error loading presentations:', error);
            alert('Error loading presentations.');
        });
}

function openPresentation(fileName) {
    fetch(`/api/powerpoint/open_presentation/${encodeURIComponent(fileName)}/`, { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('loaded successfully')) {
                currentPresentation = fileName;
                document.querySelectorAll('.file-item').forEach(item => {
                    item.classList.toggle('selected', item.innerHTML === fileName);
                });
                document.getElementById('presentation-list-container').style.display = 'none';
                document.getElementById('powerpoint-controls').style.display = 'flex';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening presentation:', error);
            alert('Error opening presentation');
        });
}

function startPresentation() {
    if (!currentPresentation) {
        alert('No presentation selected.');
        return;
    }
    fetch('/api/powerpoint/start/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Presentation started')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error starting presentation:', error);
            alert('Error starting presentation');
        });
}

function endPresentation() {
    if (!currentPresentation) {
        alert('No presentation running.');
        return;
    }
    fetch('/api/powerpoint/end/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Presentation ended')) {
                currentPresentation = null;
                document.querySelectorAll('.file-item').forEach(item => item.classList.remove('selected'));
                document.getElementById('presentation-list-container').style.display = 'block';
                document.getElementById('powerpoint-controls').style.display = 'none';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error ending presentation:', error);
            alert('Error ending presentation');
        });
}

function nextSlide() {
    fetch('/api/powerpoint/next/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Moved to next slide')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error moving to the next slide:', error);
            alert('Error moving to the next slide');
        });
}

function prevSlide() {
    fetch('/api/powerpoint/prev/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('Moved to previous slide')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error moving to the previous slide:', error);
            alert('Error moving to the previous slide');
        });
}


function openTeams() {
    hideAllModules();
    showContacts();
    fetch('/api/teams/open/', {
        method: 'GET'
    })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                document.getElementById('file-explorer').style.display = 'none';
                document.getElementById('teams').style.display = 'flex';
                document.getElementById('partition-selection').style.display = 'none';
            } else {
                alert('Error opening Microsoft Teams');
            }
        })
        .catch(error => {
            console.error('Error opening Microsoft Teams:', error);
            alert('Error opening Microsoft Teams');
        });
}

function showFileExplorer() {
    hideAllModules();
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'none';
    document.getElementById('partition-selection').style.display = 'flex';
}

function selectPartition(partition) {
    document.getElementById('partition-selection').style.display = 'none';
    document.getElementById('file-explorer').style.display = 'flex';
    currentPath = partition;
    rootPath = partition;
    history = [partition];
    historyIndex = 0;
    fetchFileTree(partition);
}

function hideAllModules() {
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'none';
    document.getElementById('powerpoint').style.display = 'none';
    document.getElementById('partition-selection').style.display = 'none';
}

function fetchFileTree(path) {
    fetch(`/api/file_explorer/list/?path=${encodeURIComponent(path)}`)
        .then(response => response.json())
        .then(data => {
            displayFileTree(data);
            currentPath = path;
            updateNavigationButtons();
        });
}

function displayFileTree(fileTree) {
    const fileTreeContainer = document.getElementById('file-tree');
    fileTreeContainer.innerHTML = '';

    fileTree.forEach(file => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `<i class="fas ${getFileIcon(file.extension)}"></i> ${file.name}`;
        item.onclick = (event) => handleFileClick(event, file);
        item.ondblclick = (event) => handleFileDoubleClick(event, file);
        fileTreeContainer.appendChild(item);
    });
}

function getFileIcon(extension) {
    switch (extension) {
        case 'folder': return 'fa-folder';
        case 'txt': return 'fa-file-alt';
        case 'pdf': return 'fa-file-pdf';
        case 'image': return 'fa-file-image';
        default: return 'fa-file';
    }
}

function handleFileClick(event, file) {
    clearSelection();
    event.currentTarget.classList.add('selected');
    selectedFilePath = file.path;

    if (file.extension === 'folder') {
        selectedFolder = file.path;
    } else {
        previewFile(file.path);
    }
}

function handleFileDoubleClick(event, file) {
    if (file.extension === 'folder') {
        history = history.slice(0, historyIndex + 1);
        history.push(file.path);
        historyIndex++;
        fetchFileTree(file.path);
        selectedFolder = null;
        clearPreview();
    }
}

function clearSelection() {
    document.querySelectorAll('.file-item').forEach(item => {
        item.classList.remove('selected');
    });
}

function previewFile(path) {
    const filePreview = document.getElementById('file-preview');
    fetch(`/api/file_explorer/download/?path=${encodeURIComponent(path)}`)
        .then(response => response.blob())
        .then(blob => {
            const url = URL.createObjectURL(blob);
            const fileType = blob.type;
            const fileExtension = path.split('.').pop().toLowerCase();

            let content = '';
            if (fileType.startsWith('image/')) {
                content = `<img src="${url}" alt="Preview">`;
            } else if (fileExtension === 'txt') {
                const reader = new FileReader();
                reader.onload = function(event) {
                    filePreview.innerHTML = `<button class="close-btn" onclick="clearPreview()">&times;</button><pre>${event.target.result}</pre>`;
                };
                reader.readAsText(blob);
                return;
            } else if (fileExtension === 'pdf') {
                content = `<iframe src="${url}" type="application/pdf"></iframe>`;
            } else {
                content = `<a href="${url}" download>Download File</a>`;
            }
            filePreview.innerHTML = `<button class="close-btn" onclick="clearPreview()">&times;</button>${content}`;
        })
        .catch(error => {
            console.error('There was a problem with the fetch operation:', error);
            filePreview.innerHTML = '<p>Unable to preview the file.</p>';
        });
}

function clearPreview() {
    const filePreview = document.getElementById('file-preview');
    filePreview.innerHTML = '';
}

function goBack() {
    if (historyIndex > 0) {
        historyIndex--;
        fetchFileTree(history[historyIndex]);
    } else {
        goHome();
    }
}

function goForward() {
    if (selectedFolder) {
        history = history.slice(0, historyIndex + 1);
        history.push(selectedFolder);
        historyIndex++;
        fetchFileTree(selectedFolder);
        selectedFolder = null;
    }
}

function refresh() {
    fetchFileTree(currentPath);
}

function goHome() {
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'none';
    document.getElementById('partition-selection').style.display = 'flex';
    clearPreview();
}

function updateNavigationButtons() {
    const backButton = document.getElementById('back-button');
    backButton.disabled = (historyIndex <= 0);
}

function showContacts() {
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'flex';
    document.getElementById('partition-selection').style.display = 'none';
}
function showPPTControls(){
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'none';
    document.getElementById('partition-selection').style.display = 'none';
    document.getElementById('powerpoint').style.display = 'flex';
}

function fetchContacts() {
    fetch('/api/teams/contacts/')
        .then(response => response.json())
        .then(data => {
            displayContacts(data);
        })
        .catch(error => {
            console.error('Error fetching contacts:', error);
            document.getElementById('contacts-container').innerHTML = '<p>Error loading contacts.</p>';
        });
}

function displayContacts(contacts) {
    const container = document.getElementById('contacts-container');
    container.innerHTML = '';

    if (!contacts.length) {
        container.innerHTML = '<p>No contacts found.</p>';
        return;
    }

    contacts.forEach(contact => {
        const item = document.createElement('div');
        item.className = 'contact-item';
        item.innerHTML = `
            <img src="${contact.photoUrl || 'https://via.placeholder.com/40'}" alt="${contact.displayName}">
            <span>${contact.displayName}</span>
        `;
        item.onclick = () => handleContactClick(contact);
        container.appendChild(item);
    });

    container.style.display = 'block';
}

function handleContactClick(contact) {
    // Clear the previously selected contact
    const contactItems = document.querySelectorAll('.contact-item');
    contactItems.forEach(item => item.classList.remove('selected'));

    // Select the new contact
    const selectedItem = Array.from(contactItems).find(item => item.textContent.includes(contact.displayName));
    if (selectedItem) {
        selectedItem.classList.add('selected');
    }

    selectedContactId = contact.id;
    document.getElementById('call-button').style.display = 'flex'; // Show the call button
}

function handleCall() {
    if (selectedContactId) {
        fetch(`/api/teams/call/${selectedContactId}/`, {
            method: 'POST'
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                alert('Calling ' + selectedContactId);
            } else {
                alert('Error initiating call');
            }
        })
        .catch(error => {
            console.error('Error initiating call:', error);
            alert('Error initiating call');
        });
    } else {
        alert('No contact selected');
    }
}

