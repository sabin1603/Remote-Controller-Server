function hideAllModules() {
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'none';
    document.getElementById('powerpoint').style.display = 'none';
    document.getElementById('partition-selection').style.display = 'none';
    document.getElementById('excel').style.display = 'none';
    document.getElementById('word').style.display = 'none';
    document.getElementById('pdf-reader').style.display = 'none';
    document.getElementById('chrome').style.display = 'none';
    document.getElementById('skype').style.display = 'none';
}

// Helper functions for consistency
function handleApiResponse(response) {
    return response.text().then(text => {
        console.log("Response received: ", text);
        const data = JSON.parse(text);
        if (!response.ok) {
            throw new Error(data.message || 'Unknown error');
        }
        return data;
    });
}

function handleApiError(error) {
    console.error('API Error:', error);
    showAlert('An error occurred: ' + error.message);
}

function showAlert(message) {
    alert(message);
}