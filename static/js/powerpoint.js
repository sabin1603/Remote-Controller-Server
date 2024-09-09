let currentPresentation = null;

function openPowerpoint() {
    hideAllModules();
    fetch('/api/powerpoint/controls/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('powerpoint').style.display = 'flex';
                document.getElementById('powerpoint-controls').style.display = 'flex';
            } else {
                showAlert('Error loading PowerPoint controls.');
            }
        })
        .catch(handleApiError);
}

function openPresentation(filePath) {
    fetch(`/api/powerpoint/open_presentation/${encodeURIComponent(filePath)}/`, { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message === "Presentation loaded successfully.") {
                currentPresentation = filePath;
                openPowerpoint();
                showAlert('Presentation opened.');
            } else {
                showAlert('Error opening Presentation: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function closePresentation() {
    fetch('/api/powerpoint/close/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes("Presentation closed successfully.")) {
                currentPresentation = null;
                showAlert(data.message);
            } else {
                showAlert('Error: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function startPresentation() {
    fetch('/api/powerpoint/start/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Presentation started')) {
                showAlert(data.message);
            } else {
                showAlert('Error: ' + data.message);
            }
        })
        .catch(handleApiError);
}

function endPresentation() {
    if (!currentPresentation) {
        showAlert('No presentation running.');
        return;
    }
    fetch('/api/powerpoint/end/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Presentation ended')) {
                currentPresentation = null;
                document.querySelectorAll('.file-item').forEach(item => item.classList.remove('selected'));
                document.getElementById('presentation-list-container').style.display = 'block';
                document.getElementById('powerpoint-controls').style.display = 'none';
            } else {
                showAlert(data.message);
            }
        })
        .catch(handleApiError);
}

function nextSlide() {
    fetch('/api/powerpoint/next/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Moved to next slide')) {
                showAlert(data.message);
            } else {
                showAlert(data.message);
            }
        })
        .catch(handleApiError);
}

function prevSlide() {
    fetch('/api/powerpoint/prev/', { method: 'GET' })
        .then(handleApiResponse)
        .then(data => {
            if (data.message.includes('Moved to previous slide')) {
                showAlert(data.message);
            } else {
                showAlert(data.message);
            }
        })
        .catch(handleApiError);
}

