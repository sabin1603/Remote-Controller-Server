let currentPresentation = null;
setInterval(checkCurrentPresentation, 5000);
function checkCurrentPresentation() {
    fetch('/api/powerpoint/current_presentation/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.current_presentation) {
                currentPresentation = data.current_presentation;
                console.log('Current presentation:', currentPresentation);
            } else {
                currentPresentation = null;
                console.log('No presentation is currently open.');
            }
        })
        .catch(error => {
            console.error('Error fetching current presentation:', error);
        });
}

function openPowerpoint() {
    hideAllModules();
    fetch('/api/powerpoint/controls/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                document.getElementById('powerpoint').style.display = 'flex';
                document.getElementById('powerpoint-controls').style.display = 'flex';  // Ensure controls are visible
            } else {
                alert('Error loading PowerPoint controls.');
            }
        })
        .catch(error => {
            console.error('Error loading PowerPoint controls:', error);
            alert('Error loading PowerPoint controls.');
        });
}

function startPresentation() {
    checkCurrentPresentation();
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
                alert('Error: ' + data.message);
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
        .then(response => response.text())  // Get response as text
        .then(text => {
            try {
                const data = JSON.parse(text);  // Try parsing as JSON
                if (data.message.includes('Moved to next slide')) {
                    console.log(data.message);
                } else {
                    alert(data.message);
                }
            } catch (error) {
                console.error('Error parsing JSON response:', text);  // Log the full response text
                alert('Error moving to the next slide');
            }
        })
        .catch(error => {
            console.error('Error moving to the next slide:', error);
            alert('Error moving to the next slide');
        });
}

function prevSlide() {
    fetch('/api/powerpoint/prev/', { method: 'GET' })
        .then(response => response.text())  // Get response as text
        .then(text => {
            try {
                const data = JSON.parse(text);  // Try parsing as JSON
                if (data.message.includes('Moved to previous slide')) {
                    console.log(data.message);
                } else {
                    alert(data.message);
                }
            } catch (error) {
                console.error('Error parsing JSON response:', text);  // Log the full response text
                alert('Error moving to the previous slide');
            }
        })
        .catch(error => {
            console.error('Error moving to the previous slide:', error);
            alert('Error moving to the previous slide');
        });
}