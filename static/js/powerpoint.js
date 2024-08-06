let currentPresentation = null;

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