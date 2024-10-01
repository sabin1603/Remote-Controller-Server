# Remote Controller Server

Remote Controller Server is a web-based application that allows users to remotely control various file types and applications on a server PC. The app provides a seamless user interface for accessing, previewing, and managing files and apps remotely.

## How It Works

### Overview
- **Modules**: The app includes modules for File Explorer, Teams, PowerPoint, Excel, Word, PDF reader (Adobe Acrobat), Chrome and Skype.
- **File Explorer**: Browse through files on the server PC. Supports previewing `.txt`, `.pdf`, and image files. Other file types like PowerPoint, Excel, Word, and PDF can be opened directly in their respective modules.
- **File-Specific Modules**: Once a file is opened in its specific module, the user can control it remotely using functions like `bring to front`, `zoom`, `scroll`, `navigate slides`, `save`, `print`, and `toggle read mode`.
- **Teams Module**: Opens the Teams app on the server PC. Due to licensing restrictions, it uses mock data for fake contacts stored in a JSON file.
- **Chrome Module**: Provides controls like `home`, `next tab`, `previous tab`, `next page`, `previous page`, `add tab`, `close page`, `scroll`, `zoom`, and a `navigation mode` to search queries or URLs directly in Chrome on the server PC.
- **Skype Module**: Uses environment variables on the server PC to store Skype account credentials. Users can `select contacts`, `send messages`, and `make calls`.
  
## Technology Stack

### Backend Tech Stack

- **Python & Django**: Developed the API server and core application logic, enabling remote control of various applications.
- **Threading**: Managed concurrent operations for different file types, ensuring responsive interactions.
- **OOP Principles**: Designed the backend using object-oriented programming for modular and maintainable code.

### Frontend Tech Stack

- **HTML5 & CSS3**: Structured the layout and styled the interface with a responsive design, including custom elements like `buttons` and `navigation bars`.
- **JavaScript**: Implemented interactive functionalities, including application control and event handling for File Explorer, Office apps, and communication tools.
- **Libraries**:
  - **`Font Awesome`**: Added intuitive icons for a user-friendly interface.
- **Django Static Files**: Integrated static assets (CSS, JS) using Django's static files system.


## Dependencies

The project uses the following dependencies:

- **Django**: Web framework for building the backend.
- **aiohttp, aiosignal**: Asynchronous HTTP client/server framework.
- **beautifulsoup4**: For parsing HTML and XML.
- **comtypes, pywinauto**: For interacting with COM objects and automating GUI interactions on Windows.
- **msgraph-core, msgraph-sdk**: Microsoft Graph API SDK for Python.
- **PyAutoGUI, PyGetWindow, pyperclip**: For GUI automation and clipboard management.
- **selenium**: For controlling web browsers programmatically.
- **SkPy**: For interacting with Skype.
- **openpyxl, xlwings**: For working with Excel files.
- **Other utilities**: Includes libraries like `requests`, `pyjwt`, `pendulum`, `pillow`, and more for various functionalities.

For the full list of dependencies, see the [requirements.txt](./requirements.txt) file.

## Motivation

This project was developed to provide a remote control solution for managing and interacting with various file types and applications on a server PC. The motivation behind the project was to offer a centralized interface for remotely controlling apps like Office applications, Chrome, Adobe Acrobat, and Skype, simplifying file and app management tasks for users who need to operate on a server PC without direct access.

---

https://youtu.be/fsRf4nyZC7g

**Notes**:
 - The Teams module is limited to opening the Teams app due to licensing restrictions. The Skype module uses environment variables for credentials, ensuring secure handling of user information.
 - The web app has been tested on both mobile devices and PCs, including iPhones and iPads, and is fully functional across all platforms. 
