let currentPath = '/';
let rootPath = '/';
let history = [];
let historyIndex = -1;
let selectedFolder = null;
let selectedFilePath = null;

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
    } else {
        // Open the file using the server API
        if (file.path.endsWith('.xlsx')) {
            openWorkbook(file.path);
        } else if (file.path.endsWith('.pptx')) {
            openPresentation(file.path);
        } else if (file.path.endsWith('.docx')) {
            openDocument(file.path);
        } else if (file.path.endsWith('.pdf')) {
            openPdf(file.path);
        } else {
            console.error("File type unsupported.");
            alert("File type unsupported");
        }
    }
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

function clearSelection() {
    document.querySelectorAll('.file-item').forEach(item => {
        item.classList.remove('selected');
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