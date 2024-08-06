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

function showContacts() {
    document.getElementById('file-explorer').style.display = 'none';
    document.getElementById('teams').style.display = 'flex';
    document.getElementById('partition-selection').style.display = 'none';
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