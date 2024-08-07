// Function to handle opening Chrome
function openChrome() {
    hideAllModules();
    fetch('/api/chrome/open_chrome/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('opened successfully')) {
                console.log(data.message);
                // Hide other modules
                hideAllModules();
                // Show the Chrome controls
                document.getElementById('chrome').style.display = 'flex';
                document.getElementById('chrome-controls').style.display = 'flex';
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening Chrome:', error);
            alert('Error opening Chrome');
        });
}

// Add functions for Chrome controls
function goHome() {
    fetch('/api/chrome/go_home/', { method: 'POST' });
}

function goBack() {
    fetch('/api/chrome/go_back/', { method: 'POST' });
}

function goForward() {
    fetch('/api/chrome/go_forward/', { method: 'POST' });
}

function newTab() {
    fetch('/api/chrome/new_tab/', { method: 'POST' });
}

function goToLeftTab() {
    fetch('/api/chrome/go_to_left_tab/', { method: 'POST' });
}

function goToRightTab() {
    fetch('/api/chrome/go_to_right_tab/', { method: 'POST' });
}

function closeCurrentTab() {
    fetch('/api/chrome/close_current_tab/', { method: 'POST' });
}

function closeChrome() {
    fetch('/api/chrome/close_chrome/', { method: 'POST' });
}

function zoomIn() {
    fetch('/api/chrome/zoom_in/', { method: 'POST' });
}

function zoomOut() {
    fetch('/api/chrome/zoom_out/', { method: 'POST' });
}

function scrollUpChrome() {
    fetch('/api/chrome/scroll_up/', { method: 'POST' });
}

function scrollDownChrome() {
    fetch('/api/chrome/scroll_down/', { method: 'POST' });
}

function toggleNavigationMode() {
    let navModeContainer = document.getElementById('navigation-mode');
    let chromeControls = document.getElementById('chrome-controls');

    if (navModeContainer.style.display === 'none') {
        navModeContainer.style.display = 'flex';
        chromeControls.style.display = 'none';
    } else {
        navModeContainer.style.display = 'none';
        chromeControls.style.display = 'flex';
    }
}

// Navigation mode functions
function enterNavigationMode() {
    let query = document.getElementById('search-box').value;
    fetch('/api/chrome/navigate/', {
        method: 'POST',
        body: JSON.stringify({ query: query }),
        headers: { 'Content-Type': 'application/json' }
    });
}

function navigateUp() {
    fetch('/api/chrome/navigate_up/', { method: 'POST' });
}

function navigateDown() {
    fetch('/api/chrome/navigate_down/', { method: 'POST' });
}