// chrome.js

function openChrome() {
    fetch('/api/chrome/open_chrome/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('opened successfully')) {
                console.log(data.message);
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening Chrome:', error);
            alert('Error opening Chrome');
        });
}

function goHome() {
    fetch('/api/chrome/home/', { method: 'GET' })
        .then(response => response.json())
        .then(data => alert(data.message))
        .catch(error => console.error('Error:', error));
}

function goBack() {
    fetch('/api/chrome/back/', { method: 'GET' })
        .then(response => response.json())
        .then(data => alert(data.message))
        .catch(error => console.error('Error:', error));
}

function goForward() {
    fetch('/api/chrome/forward/', { method: 'GET' })
        .then(response => response.json())
        .then(data => alert(data.message))
        .catch(error => console.error('Error:', error));
}

function closeChrome() {
    fetch('/api/chrome/close/', { method: 'GET' })
        .then(response => response.json())
        .then(data => alert(data.message))
        .catch(error => console.error('Error:', error));
}

// Add similar functions for zooming, scrolling, and tab management
