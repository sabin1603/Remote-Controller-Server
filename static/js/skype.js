function openSkype() {
    hideAllModules();
    fetch('/api/skype/open/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.status === 'success') {
                console.log(data.message);
                // Hide other modules
                hideAllModules();
                // Show the Skype controls
                document.getElementById('skype').style.display = 'flex';
                document.getElementById('skype-controls').style.display = 'flex';
                fetchSkypeContacts(); // Fetch contacts after opening Skype
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error opening Skype:', error);
            alert('Error opening Skype');
        });
}

function closeSkype() {
    fetch('/api/skype/close/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.message.includes('closed successfully')) {
                console.log(data.message);
                hideAllModules();
            } else {
                alert(data.message);
            }
        })
        .catch(error => {
            console.error('Error closing Skype:', error);
            alert('Error closing Skype');
        });
}

function fetchSkypeContacts() {
    fetch('/api/skype/get_contacts/', { method: 'GET' })
        .then(response => response.json())
        .then(data => {
            if (data.contacts) {
                const contactList = document.getElementById('contact-list');
                contactList.innerHTML = ''; // Clear the previous list
                data.contacts.forEach(contact => {
                    const option = document.createElement('option');
                    option.value = contact.username;
                    option.textContent = contact.display_name;
                    contactList.appendChild(option);
                });
            } else {
                console.error('Error fetching contacts:', data.message);
            }
        })
        .catch(error => {
            console.error('Error fetching contacts:', error);
        });
}

function callSkypeUser() {
    const username = document.getElementById('contact-list').value;
    if (username) {
        fetch(`/api/skype/call/${username}/`, { method: 'GET' })
            .then(response => response.json())
            .then(data => {
                console.log(data.message);
            })
            .catch(error => {
                console.error(`Error calling ${username}:`, error);
                alert(`Error calling ${username}`);
            });
    } else {
        alert('Please select a contact to call.');
    }
}

function sendMessageToUser() {
    const username = document.getElementById('contact-list').value;
    const message = encodeURIComponent(document.getElementById('message-box').value);
    if (username && message) {
        // Fetch request to send a message
        fetch('/api/skype/message/', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({ username: username, message: decodeURIComponent(message) })
        })
            .then(response => response.json())
            .then(data => {
                console.log(data.message);
                document.getElementById('message-box').value = ''; // Clear the message box
            })
            .catch(error => {
                console.error(`Error sending message to ${username}:`, error);
                alert(`Error sending message to ${username}`);
            });
    } else {
        alert('Please select a contact and enter a message.');
    }
}




