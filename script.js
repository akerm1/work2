// Firebase Configuration
const firebaseConfig = {
    apiKey: "AIzaSyDznYcUtQWRD7QqYBDr1QupUMfVqZnfGEE",
    authDomain: "my-work-82778.firebaseapp.com",
    projectId: "my-work-82778",
    storageBucket: "my-work-82778.appspot.com",
    messagingSenderId: "1070444118182",
    appId: "1:1070444118182:web:bae373255bd124d3a2b467"
};

// Initialize Firebase safely
let db;
try {
    firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();
    console.log("✅ Firebase initialized successfully");
} catch (error) {
    console.error("❌ Firebase initialization failed:", error);
    showError("Failed to connect to Firebase. Please check your configuration.");
}

// Global variables
let currentBulkClient = null;
let accounts = [];
let clients = [];
let filteredAccounts = [];
let isSearchActive = false;
let expandedClients = new Set(); // Track expanded client sections

// Global error fallback
window.addEventListener("error", (event) => {
    console.error("❌ Runtime Error:", event.error);
    showError("An unexpected error occurred. Check console for details.");
});

// Helper function to show error in UI
function showError(message) {
    let errorBox = document.getElementById("errorBox");
    if (!errorBox) {
        errorBox = document.createElement("div");
        errorBox.id = "errorBox";
        errorBox.style.position = "fixed";
        errorBox.style.top = "20px";
        errorBox.style.left = "50%";
        errorBox.style.transform = "translateX(-50%)";
        errorBox.style.background = "#ff4444";
        errorBox.style.color = "white";
        errorBox.style.padding = "12px 20px";
        errorBox.style.borderRadius = "8px";
        errorBox.style.zIndex = "9999";
        errorBox.style.fontFamily = "Arial, sans-serif";
        errorBox.style.boxShadow = "0 2px 6px rgba(0,0,0,0.3)";
        document.body.appendChild(errorBox);
    }
    errorBox.innerText = message;
}

// Utility Functions
function formatDateForDisplay(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function formatDateForStorage(dateString) {
    if (!dateString) return getTodayFormatted();
    return dateString;
}

function getTodayFormatted() {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function getTomorrowFormatted() {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const year = tomorrow.getFullYear();
    const month = String(tomorrow.getMonth() + 1).padStart(2, '0');
    const day = String(tomorrow.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function showLoading() {
    document.getElementById('loadingSpinner').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loadingSpinner').style.display = 'none';
}

function showMessage(message, type = 'success') {
    const messageEl = document.createElement('div');
    messageEl.className = `message ${type}`;
    messageEl.textContent = message;
    document.body.appendChild(messageEl);
    
    setTimeout(() => {
        messageEl.remove();
    }, 4000);
}

function copyToClipboard(text) {
    copyToClipboardSafe(text, 'Emails copied to clipboard');
}

// Order Number Functions
function renderOrderNumbers(orderNumbers) {
    if (!orderNumbers) orderNumbers = ['', '', '', '', ''];
    
    return `<div class="order-numbers">
        ${orderNumbers.map((order, index) => {
            if (order && order.trim()) {
                return `
                    <div class="order-number">
                        <span class="order-text">${escapeHtml(order)}</span>
                        <button class="order-copy-btn" type="button" data-copy="${escapeHtml(order)}" title="Copy ${escapeHtml(order)}">
                            📋
                        </button>
                    </div>
                `;
            } else {
                return `
                    <div class="order-number empty">
                        <span class="order-text">Order ${index + 1}</span>
                    </div>
                `;
            }
        }).join('')}
    </div>`;
}

function renderOrderInputs(orderNumbers, accountId) {
    if (!orderNumbers) orderNumbers = ['', '', '', '', ''];
    
    return `<div class="order-inputs">
        ${orderNumbers.map((order, index) => `
            <input type="text" 
                   value="${escapeHtml(order)}" 
                   placeholder="Order ${index + 1}"
                   data-account="${escapeHtml(accountId)}"
                   data-index="${index}"
                   class="order-input"
                   maxlength="50">
        `).join('')}
    </div>`;
}

// Security helper function
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, function(m) { return map[m]; });
}

// Safe copy function with fallback
function copyToClipboardSafe(text, successMessage = 'Copied to clipboard') {
    // Modern clipboard API
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text)
            .then(() => showMessage(successMessage))
            .catch(err => {
                console.warn('Clipboard API failed, using fallback');
                fallbackCopy(text, successMessage);
            });
    } else {
        // Fallback for older browsers or non-HTTPS
        fallbackCopy(text, successMessage);
    }
}

function fallbackCopy(text, successMessage) {
    try {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        
        const successful = document.execCommand('copy');
        document.body.removeChild(textArea);
        
        if (successful) {
            showMessage(successMessage);
        } else {
            showMessage('Copy failed - please select and copy manually', 'error');
        }
    } catch (err) {
        showMessage('Copy not supported in this browser', 'error');
        console.error('Copy failed:', err);
    }
}

function searchByOrderNumber() {
    const query = document.getElementById('orderSearchInput').value.trim().toLowerCase();
    
    if (!query) {
        document.getElementById('orderResults').style.display = 'none';
        return;
    }

    const matches = accounts.filter(account => 
        account.orderNumbers && account.orderNumbers.some(order => 
            order.toLowerCase().includes(query)
        )
    );

    renderOrderSearchResults(matches, query);
}

function renderOrderSearchResults(matches, query) {
    const resultsDiv = document.getElementById('orderResults');
    
    if (matches.length === 0) {
        resultsDiv.innerHTML = `<div class="no-data">No accounts found with order number containing "${query}"</div>`;
        resultsDiv.style.display = 'block';
        return;
    }

    const resultsHTML = matches.map(account => {
        const matchingOrders = account.orderNumbers.filter(order => 
            order.toLowerCase().includes(query)
        ).join(', ');

        return `
            <div class="order-match">
                <div class="order-match-header">
                    ${escapeHtml(account.client)} - ${escapeHtml(account.email)}
                </div>
                <div class="order-match-details">
                    Matching Orders: <strong>${escapeHtml(matchingOrders)}</strong><br>
                    All Orders: ${escapeHtml(account.orderNumbers.filter(o => o).join(', ') || 'None')}<br>
                    Expiration: ${formatDateForDisplay(account.date)}
                </div>
            </div>
        `;
    }).join('');

    resultsDiv.innerHTML = `
        <h4>Found ${matches.length} account${matches.length !== 1 ? 's' : ''} with matching order numbers:</h4>
        ${resultsHTML}
    `;
    resultsDiv.style.display = 'block';
}

// Enhanced security for dynamic content
function sanitizeForHTML(str) {
    return escapeHtml(String(str || ''));
}

function createSecureElement(tag, attributes = {}, textContent = '') {
    const element = document.createElement(tag);
    
    // Set attributes securely
    Object.keys(attributes).forEach(key => {
        if (key === 'onclick') {
            // Avoid inline onclick handlers
            return;
        }
        element.setAttribute(key, sanitizeForHTML(attributes[key]));
    });
    
    if (textContent) {
        element.textContent = textContent;
    }
    
    return element;
}

function clearOrderSearch() {
    document.getElementById('orderSearchInput').value = '';
    document.getElementById('orderResults').style.display = 'none';
}

// Event delegation for order copy buttons
function handleOrderCopy(event) {
    if (event.target.classList.contains('order-copy-btn')) {
        event.preventDefault();
        event.stopPropagation();
        const orderText = event.target.getAttribute('data-copy');
        if (orderText && orderText.trim()) {
            copyToClipboardSafe(orderText.trim(), `Order "${orderText}" copied to clipboard`);
        }
    }
}

// Firebase Functions
async function loadAccounts() {
    if (!db) {
        showError("Firestore not available. Accounts cannot be loaded.");
        return;
    }

    try {
        showLoading();
        const snapshot = await db.collection('accounts').get();
        accounts = [];
        
        if (snapshot.empty) {
            console.warn("⚠️ No accounts found in Firestore.");
            clients = [];
            renderAccounts();
            renderExpiringAccounts();
            hideLoading();
            return;
        }

        snapshot.forEach(doc => {
            const data = doc.data();
            console.log("🔥 Account:", doc.id, data);
            
            // Handle both old accounts (without orderNumbers) and new accounts (with orderNumbers)
            const account = {
                id: doc.id,
                client: data.client,
                email: data.email,
                date: data.date,
                orderNumbers: data.orderNumbers || ['', '', '', '', ''] // Default empty order numbers for existing accounts
            };
            
            accounts.push(account);
        });
        
        clients = [...new Set(accounts.map(acc => acc.client))].filter(Boolean);
        
        console.log('Loaded accounts:', accounts.map(acc => acc.email));
        
        renderAccounts();
        renderExpiringAccounts();
        hideLoading();
    } catch (error) {
        console.error('Error loading accounts:', error);
        showMessage('Error loading accounts', 'error');
        hideLoading();
    }
}

async function saveAccount(accountData) {
    try {
        const docRef = await db.collection('accounts').add({
            client: accountData.client,
            email: accountData.email,
            date: formatDateForStorage(accountData.date),
            orderNumbers: accountData.orderNumbers || ['', '', '', '', '']
        });
        return docRef.id;
    } catch (error) {
        console.error('Error saving account:', error);
        throw error;
    }
}

async function updateAccount(id, accountData) {
    try {
        await db.collection('accounts').doc(id).update({
            client: accountData.client,
            email: accountData.email,
            date: formatDateForStorage(accountData.date),
            orderNumbers: accountData.orderNumbers || ['', '', '', '', '']
        });
    } catch (error) {
        console.error('Error updating account:', error);
        throw error;
    }
}

async function deleteAccount(id) {
    try {
        await db.collection('accounts').doc(id).delete();
    } catch (error) {
        console.error('Error deleting account:', error);
        throw error;
    }
}

async function bulkSaveAccounts(accountsData) {
    try {
        const batch = db.batch();
        
        accountsData.forEach(accountData => {
            const docRef = db.collection('accounts').doc();
            batch.set(docRef, {
                client: accountData.client,
                email: accountData.email,
                date: formatDateForStorage(accountData.date),
                orderNumbers: accountData.orderNumbers || ['', '', '', '', '']
            });
        });
        
        await batch.commit();
    } catch (error) {
        console.error('Error bulk saving accounts:', error);
        throw error;
    }
}

// Search Functions
function performSearch() {
    const emailInput = document.getElementById('emailSearchInput').value.trim();
    const dateQuery = document.getElementById('dateSearchInput').value;
    const clientQuery = document.getElementById('clientSearchInput').value.trim().toLowerCase();
    
    if (!emailInput && !dateQuery && !clientQuery) {
        clearSearch();
        return;
    }
    
    // Split input by commas, newlines, or separators, and normalize
    const emailQueries = emailInput.split(/[\n,\s-]+/)
        .map(email => email.trim().toLowerCase())
        .filter(email => {
            // Validate email format
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            return email && emailRegex.test(email);
        });
    
    console.log('Search email queries:', emailQueries);
    console.log('Database emails:', accounts.map(acc => acc.email.toLowerCase()));
    
    // First, try to find exact email matches
    filteredAccounts = accounts.filter(account => {
        const accountEmail = account.email.toLowerCase().trim();
        const exactEmailMatch = emailQueries.some(query => accountEmail === query);
        const partialEmailMatch = emailQueries.some(query => accountEmail.includes(query));
        const emailMatch = !emailInput || exactEmailMatch || (emailQueries.length > 0 && partialEmailMatch);
        const dateMatch = !dateQuery || account.date === dateQuery;
        const clientMatch = !clientQuery || account.client.toLowerCase().includes(clientQuery);
        
        return emailMatch && dateMatch && clientMatch;
    });
    
    console.log('Filtered accounts:', filteredAccounts);
    
    isSearchActive = true;
    renderSearchResults();
    
    if (filteredAccounts.length === 0) {
        showMessage('No accounts found matching your search criteria.', 'error');
    }
}

function clearSearch() {
    document.getElementById('emailSearchInput').value = '';
    document.getElementById('dateSearchInput').value = '';
    document.getElementById('clientSearchInput').value = '';
    document.getElementById('searchResults').style.display = 'none';
    isSearchActive = false;
    filteredAccounts = [];
}

function renderSearchResults() {
    const resultsContainer = document.getElementById('searchResults');
    const resultsBody = document.getElementById('searchResultsBody');
    const resultCount = document.getElementById('resultCount');
    
    if (filteredAccounts.length === 0) {
        resultsBody.innerHTML = '<tr><td colspan="5" class="no-data">No accounts found matching your search criteria.</td></tr>';
        resultCount.textContent = '0 results found';
        resultsContainer.style.display = 'block';
        return;
    }
    
    // Group accounts by client
    const accountsByClient = {};
    filteredAccounts.forEach(account => {
        if (!accountsByClient[account.client]) {
            accountsByClient[account.client] = [];
        }
        accountsByClient[account.client].push(account);
    });
    
    // Render search results grouped by client
    resultsBody.innerHTML = Object.keys(accountsByClient).sort().map(client => `
        <tr>
            <td colspan="5" class="client-header">
                <div class="client-name">${client} (${accountsByClient[client].length} account${accountsByClient[client].length !== 1 ? 's' : ''})</div>
                <button class="btn btn-copy btn-small" onclick="copyToClipboard('${accountsByClient[client].map(acc => acc.email).join(', ')}')">
                    Copy All Emails
                </button>
            </td>
        </tr>
        ${accountsByClient[client].map(account => `
            <tr class="account-row">
                <td>${account.client}</td>
                <td>${account.email}</td>
                <td>${formatDateForDisplay(account.date)}</td>
                <td>${renderOrderNumbers(account.orderNumbers)}</td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn-warning btn-small" onclick="editAccount('${account.id}')">
                            Edit
                        </button>
                        <button class="btn btn-danger btn-small" onclick="confirmDeleteAccount('${account.id}')">
                            Delete
                        </button>
                        <button class="btn btn-copy btn-small" onclick="copyToClipboard('${account.email}')">
                            Copy
                        </button>
                    </div>
                </td>
            </tr>
        `).join('')}
    `).join('');
    
    resultCount.textContent = `${filteredAccounts.length} result${filteredAccounts.length !== 1 ? 's' : ''} found`;
    resultsContainer.style.display = 'block';
}

// Export Functions
function exportToCSV() {
    const dataToExport = isSearchActive ? filteredAccounts : accounts;
    
    if (dataToExport.length === 0) {
        showMessage('No accounts to export', 'error');
        return;
    }
    
    const headers = ['Client', 'Email', 'Expiration Date', 'Order 1', 'Order 2', 'Order 3', 'Order 4', 'Order 5'];
    const csvContent = [
        headers.join(','),
        ...dataToExport.map(account => [
            `"${account.client}"`,
            `"${account.email}"`,
            `"${formatDateForDisplay(account.date)}"`,
            `"${account.orderNumbers ? account.orderNumbers[0] || '' : ''}"`,
            `"${account.orderNumbers ? account.orderNumbers[1] || '' : ''}"`,
            `"${account.orderNumbers ? account.orderNumbers[2] || '' : ''}"`,
            `"${account.orderNumbers ? account.orderNumbers[3] || '' : ''}"`,
            `"${account.orderNumbers ? account.orderNumbers[4] || '' : ''}"`
        ].join(','))
    ].join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `email-accounts-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    const exportType = isSearchActive ? 'filtered' : 'all';
    showMessage(`CSV file with ${exportType} accounts downloaded successfully`);
}

function exportToExcel() {
    const dataToExport = isSearchActive ? filteredAccounts : accounts;
    
    if (dataToExport.length === 0) {
        showMessage('No accounts to export', 'error');
        return;
    }
    
    const excelContent = `<?xml version="1.0"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
        <Worksheet ss:Name="Email Accounts">
            <Table>
                <Row>
                    <Cell><Data ss:Type="String">Client</Data></Cell>
                    <Cell><Data ss:Type="String">Email</Data></Cell>
                    <Cell><Data ss:Type="String">Expiration Date</Data></Cell>
                    <Cell><Data ss:Type="String">Order 1</Data></Cell>
                    <Cell><Data ss:Type="String">Order 2</Data></Cell>
                    <Cell><Data ss:Type="String">Order 3</Data></Cell>
                    <Cell><Data ss:Type="String">Order 4</Data></Cell>
                    <Cell><Data ss:Type="String">Order 5</Data></Cell>
                </Row>
                ${dataToExport.map(account => `
                    <Row>
                        <Cell><Data ss:Type="String">${account.client}</Data></Cell>
                        <Cell><Data ss:Type="String">${account.email}</Data></Cell>
                        <Cell><Data ss:Type="String">${formatDateForDisplay(account.date)}</Data></Cell>
                        <Cell><Data ss:Type="String">${account.orderNumbers ? account.orderNumbers[0] || '' : ''}</Data></Cell>
                        <Cell><Data ss:Type="String">${account.orderNumbers ? account.orderNumbers[1] || '' : ''}</Data></Cell>
                        <Cell><Data ss:Type="String">${account.orderNumbers ? account.orderNumbers[2] || '' : ''}</Data></Cell>
                        <Cell><Data ss:Type="String">${account.orderNumbers ? account.orderNumbers[3] || '' : ''}</Data></Cell>
                        <Cell><Data ss:Type="String">${account.orderNumbers ? account.orderNumbers[4] || '' : ''}</Data></Cell>
                    </Row>
                `).join('')}
            </Table>
        </Worksheet>
    </Workbook>`;
    
    const blob = new Blob([excelContent], { type: 'application/vnd.ms-excel' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `email-accounts-${new Date().toISOString().split('T')[0]}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    const exportType = isSearchActive ? 'filtered' : 'all';
    showMessage(`Excel file with ${exportType} accounts downloaded successfully`);
}

function exportToJSON() {
    const dataToExport = isSearchActive ? filteredAccounts : accounts;
    
    if (dataToExport.length === 0) {
        showMessage('No accounts to export', 'error');
        return;
    }
    
    const jsonContent = JSON.stringify(dataToExport.map(account => ({
        client: account.client,
        email: account.email,
        expirationDate: formatDateForDisplay(account.date),
        rawDate: account.date,
        orderNumbers: account.orderNumbers || ['', '', '', '', '']
    })), null, 2);
    
    const blob = new Blob([jsonContent], { type: 'application/json' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `email-accounts-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    const exportType = isSearchActive ? 'filtered' : 'all';
    showMessage(`JSON file with ${exportType} accounts downloaded successfully`);
}

// Rendering Functions
function renderAccounts() {
    const container = document.getElementById('clientsContainer');
    container.innerHTML = '';
    
    const accountsByClient = {};
    accounts.forEach(account => {
        if (!accountsByClient[account.client]) {
            accountsByClient[account.client] = [];
        }
        accountsByClient[account.client].push(account);
    });
    
    Object.keys(accountsByClient).sort().forEach(client => {
        const clientSection = createClientSection(client, accountsByClient[client]);
        container.appendChild(clientSection);
    });
    
    if (Object.keys(accountsByClient).length === 0) {
        container.innerHTML = '<div class="no-data">No accounts found. Add your first client to get started!</div>';
    }
}

function createClientSection(clientName, clientAccounts) {
    const isExpanded = expandedClients.has(clientName);
    const displayAccounts = isExpanded ? clientAccounts : clientAccounts.slice(0, 5);
    const showToggle = clientAccounts.length > 5;
    
    const section = document.createElement('div');
    section.className = 'client-section';
    section.innerHTML = `
        <div class="client-header">
            <div class="client-name">${clientName} (${clientAccounts.length} account${clientAccounts.length !== 1 ? 's' : ''})</div>
            <div class="client-actions">
                <button class="btn btn-success btn-small" onclick="addNewAccount('${clientName}')">
                    + Add New Account
                </button>
                <button class="btn btn-primary btn-small" onclick="openBulkUpload('${clientName}')">
                    + Bulk Upload Emails
                </button>
                <button class="btn btn-copy btn-small" onclick="copyToClipboard('${clientAccounts.map(acc => acc.email).join(', ')}')">
                    Copy All Emails
                </button>
            </div>
        </div>
        <div class="table-container" style="max-height: ${showToggle && !isExpanded ? '300px' : 'none'}; overflow-y: ${showToggle && !isExpanded ? 'auto' : 'visible'};">
            <table>
                <thead>
                    <tr>
                        <th>Email</th>
                        <th>Expiration Date</th>
                        <th>Order Numbers</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody id="client-${clientName.replace(/\s+/g, '-')}">
                    ${displayAccounts.map(account => createAccountRow(account)).join('')}
                </tbody>
            </table>
        </div>
        ${showToggle ? `
            <div class="toggle-container" style="text-align: center; margin-top: 10px;">
                <button class="btn btn-primary btn-small" onclick="toggleClientAccounts('${clientName}')">
                    ${isExpanded ? 'Show Less' : 'Show More'}
                </button>
            </div>
        ` : ''}
    `;
    
    return section;
}

function toggleClientAccounts(clientName) {
    if (expandedClients.has(clientName)) {
        expandedClients.delete(clientName);
    } else {
        expandedClients.add(clientName);
    }
    renderAccounts();
}

function createAccountRow(account) {
    return `
        <tr class="account-row" data-id="${account.id}">
            <td class="email-cell">${account.email}</td>
            <td class="date-cell">${formatDateForDisplay(account.date)}</td>
            <td class="orders-cell">${renderOrderNumbers(account.orderNumbers)}</td>
            <td class="actions-cell">
                <div class="action-buttons">
                    <button class="btn btn-warning btn-small" onclick="editAccount('${account.id}')">
                        Edit
                    </button>
                    <button class="btn btn-danger btn-small" onclick="confirmDeleteAccount('${account.id}')">
                        Delete
                    </button>
                    <button class="btn btn-copy btn-small" onclick="copyToClipboard('${account.email}')">
                        Copy
                    </button>
                </div>
            </td>
        </tr>
    `;
}

function renderExpiringAccounts() {
    const tomorrow = getTomorrowFormatted();
    const expiringAccounts = accounts.filter(account => account.date === tomorrow);
    
    const tbody = document.getElementById('expiringTableBody');
    const noDataDiv = document.getElementById('noExpiringAccounts');
    
    if (expiringAccounts.length === 0) {
        tbody.innerHTML = '';
        noDataDiv.style.display = 'block';
    } else {
        noDataDiv.style.display = 'none';
        tbody.innerHTML = expiringAccounts.map(account => `
            <tr class="account-row">
                <td>${account.client}</td>
                <td>${account.email}</td>
                <td>${formatDateForDisplay(account.date)}</td>
                <td>${renderOrderNumbers(account.orderNumbers)}</td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn-warning btn-small" onclick="editAccount('${account.id}')">
                            Edit
                        </button>
                        <button class="btn btn-danger btn-small" onclick="confirmDeleteAccount('${account.id}')">
                            Delete
                        </button>
                        <button class="btn btn-copy btn-small" onclick="copyToClipboard('${account.email}')">
                            Copy
                        </button>
                    </div>
                </td>
            </tr>
        `).join('');
    }
}

// Account Management Functions
function addNewAccount(clientName) {
    const tbody = document.getElementById(`client-${clientName.replace(/\s+/g, '-')}`);
    const newRow = document.createElement('tr');
    newRow.className = 'account-row editing-row';
    newRow.innerHTML = `
        <td>
            <input type="email" class="new-email" placeholder="Enter email" required>
        </td>
        <td>
            <input type="date" class="new-date" value="${getTodayFormatted()}">
        </td>
        <td>
            ${renderOrderInputs(['', '', '', '', ''], 'new')}
        </td>
        <td>
            <div class="action-buttons">
                <button class="btn btn-success btn-small" onclick="saveNewAccount('${clientName}', this)">
                    Save
                </button>
                <button class="btn btn-secondary btn-small" onclick="cancelNewAccount(this)">
                    Cancel
                </button>
            </div>
        </td>
    `;
    
    tbody.appendChild(newRow);
    newRow.querySelector('.new-email').focus();
}

// Input validation functions
function validateEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

function validateDate(dateString) {
    if (!dateString) return true; // Empty dates are OK
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(dateString)) return false;
    
    const date = new Date(dateString);
    return date instanceof Date && !isNaN(date);
}

function sanitizeOrderNumber(orderNum) {
    // Remove potentially dangerous characters, keep only alphanumeric, spaces, hyphens, underscores
    return String(orderNum || '').replace(/[^a-zA-Z0-9\s\-_]/g, '').trim().substring(0, 50);
}

async function saveNewAccount(clientName, button) {
    const row = button.closest('tr');
    const email = row.querySelector('.new-email').value.trim();
    const date = row.querySelector('.new-date').value;
    const orderInputs = row.querySelectorAll('.order-input');
    const orderNumbers = Array.from(orderInputs).map(input => sanitizeOrderNumber(input.value));
    
    if (!email) {
        showMessage('Please enter an email address', 'error');
        return;
    }
    
    if (!validateEmail(email)) {
        showMessage('Please enter a valid email address', 'error');
        return;
    }
    
    if (!validateDate(date)) {
        showMessage('Please enter a valid date', 'error');
        return;
    }
    
    try {
        showLoading();
        const accountData = {
            client: sanitizeForHTML(clientName),
            email: sanitizeForHTML(email),
            date: date || getTodayFormatted(),
            orderNumbers: orderNumbers
        };
        
        const id = await saveAccount(accountData);
        accounts.push({ id, ...accountData });
        
        hideLoading();
        showMessage('Account added successfully');
        renderAccounts();
        renderExpiringAccounts();
        
        if (isSearchActive) {
            performSearch();
        }
    } catch (error) {
        hideLoading();
        showMessage('Error adding account', 'error');
        console.error('Save error:', error);
    }
}

function cancelNewAccount(button) {
    const row = button.closest('tr');
    row.remove();
}

function editAccount(id) {
    const account = accounts.find(acc => acc.id === id);
    if (!account) {
        showMessage('Account not found', 'error');
        return;
    }

    let row;

    // First, try to find the row in the SEARCH RESULTS table
    if (isSearchActive) {
        row = document.querySelector(`#searchResultsBody tr[data-id="${id}"]`);
    } else {
        // If not in search mode, find it in the main clients table
        row = document.querySelector(`tr[data-id="${id}"]`);
    }

    if (!row) {
        console.error(`Could not find row for account ID: ${id}`);
        showMessage('Could not find account to edit. Please try again.', 'error');
        return;
    }

    // Identify cells based on their position in the row
    // Search results have 5 columns: Client, Email, Date, Orders, Actions
    // Main table has 4 columns: Email, Date, Orders, Actions
    const cells = row.querySelectorAll('td');
    let emailCell, dateCell, ordersCell, actionsCell;

    if (isSearchActive) {
        // In search results: Email is 2nd cell (index 1), Date is 3rd (index 2), Orders is 4th (index 3), Actions is 5th (index 4)
        emailCell = cells[1];
        dateCell = cells[2];
        ordersCell = cells[3];
        actionsCell = cells[4];
    } else {
        // In main table: Email is 1st cell (index 0), Date is 2nd (index 1), Orders is 3rd (index 2), Actions is 4th (index 3)
        emailCell = cells[0];
        dateCell = cells[1];
        ordersCell = cells[2];
        actionsCell = cells[3];
    }

    if (!emailCell || !dateCell || !ordersCell || !actionsCell) {
        console.error('Could not identify table cells for editing');
        showMessage('Error preparing edit form', 'error');
        return;
    }

    // Create the edit form
    emailCell.innerHTML = `<input type="email" class="edit-email" value="${escapeHtml(account.email)}">`;
    dateCell.innerHTML = `<input type="date" class="edit-date" value="${account.date}">`;
    ordersCell.innerHTML = renderOrderInputs(account.orderNumbers, id);
    actionsCell.innerHTML = `
        <div class="action-buttons">
            <button class="btn btn-success btn-small" onclick="saveAccountEdit('${id}')">
                Save
            </button>
            <button class="btn btn-secondary btn-small" onclick="cancelAccountEdit('${id}')">
                Cancel
            </button>
        </div>
    `;

    row.classList.add('editing-row');
    row.querySelector('.edit-email').focus();
}

async function saveAccountEdit(id) {
    // Find the row in either search results or main table
    let row = document.querySelector(`#searchResultsBody tr[data-id="${id}"]`);
    if (!row) {
        row = document.querySelector(`tr[data-id="${id}"]`);
    }
    
    if (!row) {
        showMessage('Could not find account row to save', 'error');
        return;
    }

    const emailInput = row.querySelector('.edit-email');
    const dateInput = row.querySelector('.edit-date');
    const orderInputs = row.querySelectorAll('.order-input');

    if (!emailInput || !dateInput) {
        showMessage('Error: Edit form elements not found', 'error');
        return;
    }

    const email = emailInput.value.trim();
    const date = dateInput.value;
    const orderNumbers = Array.from(orderInputs).map(input => sanitizeOrderNumber(input.value));

    if (!email || !validateEmail(email)) {
        showMessage('Please enter a valid email address', 'error');
        return;
    }
    if (!validateDate(date)) {
        showMessage('Please enter a valid date', 'error');
        return;
    }

    try {
        showLoading();
        const account = accounts.find(acc => acc.id === id);
        const updatedData = {
            client: sanitizeForHTML(account.client),
            email: sanitizeForHTML(email),
            date: date || getTodayFormatted(),
            orderNumbers: orderNumbers
        };
        await updateAccount(id, updatedData);
        Object.assign(account, updatedData);
        hideLoading();
        showMessage('Account updated successfully');
        
        // Re-render based on current mode
        if (isSearchActive) {
            renderSearchResults(); // Stay in search results
        } else {
            renderAccounts(); // Update main view
        }
        renderExpiringAccounts(); // Always update expiring accounts
    } catch (error) {
        hideLoading();
        showMessage('Error updating account', 'error');
        console.error('Update error:', error);
    }
}

function cancelAccountEdit(id) {
    // Simply re-render the appropriate view
    if (isSearchActive) {
        renderSearchResults(); // Stay in search results
    } else {
        renderAccounts(); // Go back to main view
    }
}
function cancelAccountEdit(id) {
    // Simply re-render the search results to go back to view mode
    renderSearchResults();
}

async function deleteAccountConfirmed(id) {
    try {
        showLoading();
        await deleteAccount(id);
        
        accounts = accounts.filter(acc => acc.id !== id);
        
        hideLoading();
        showMessage('Account deleted successfully');
        renderAccounts();
        renderExpiringAccounts();
        
        if (isSearchActive) {
            performSearch();
        }
    } catch (error) {
        hideLoading();
        showMessage('Error deleting account', 'error');
    }
}

// Bulk Upload Functions
function openBulkUpload(clientName) {
    currentBulkClient = clientName;
    const modal = document.getElementById('bulkUploadModal');
    const input = document.getElementById('bulkEmailInput');
    
    input.value = '';
    modal.style.display = 'block';
    input.focus();
}

async function processBulkUpload() {
    const input = document.getElementById('bulkEmailInput').value.trim();
    
    if (!input) {
        showMessage('Please enter some emails', 'error');
        return;
    }
    
    const lines = input.split('\n').filter(line => line.trim());
    const accountsToAdd = [];
    const errors = [];
    
    lines.forEach((line, index) => {
        const trimmedLine = line.trim();
        if (!trimmedLine) return;
        
        const parts = trimmedLine.split(',').map(part => part.trim());
        const email = parts[0];
        const date = parts[1] || getTodayFormatted();
        
        // Extract order numbers from parts 2-6 (up to 5 order numbers)
        const orderNumbers = [];
        for (let i = 2; i < 7; i++) {
            orderNumbers.push(parts[i] || '');
        }
        
        // Ensure we have exactly 5 order number slots
        while (orderNumbers.length < 5) {
            orderNumbers.push('');
        }
        
        if (!email || !email.includes('@')) {
            errors.push(`Line ${index + 1}: Invalid email format`);
            return;
        }
        
        if (date && date !== getTodayFormatted()) {
            const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
            if (!dateRegex.test(date)) {
                errors.push(`Line ${index + 1}: Invalid date format (use YYYY-MM-DD)`);
                return;
            }
        }
        
        accountsToAdd.push({
            client: currentBulkClient,
            email: email,
            date: date,
            orderNumbers: orderNumbers.slice(0, 5)
        });
    });
    
    if (errors.length > 0) {
        showMessage(errors.join('\n'), 'error');
        return;
    }
    
    if (accountsToAdd.length === 0) {
        showMessage('No valid emails found', 'error');
        return;
    }
    
    try {
        showLoading();
        await bulkSaveAccounts(accountsToAdd);
        
        // Reload accounts from database to get the actual IDs
        await loadAccounts();
        
        hideLoading();
        closeBulkUploadModal();
        showMessage(`Successfully added ${accountsToAdd.length} accounts`);
    } catch (error) {
        hideLoading();
        showMessage('Error uploading accounts', 'error');
        console.error('Bulk upload error:', error);
    }
}

function closeBulkUploadModal() {
    document.getElementById('bulkUploadModal').style.display = 'none';
    currentBulkClient = null;
}

// Client Management Functions
function showAddClientForm() {
    document.getElementById('addClientBtn').style.display = 'none';
    document.getElementById('newClientForm').style.display = 'flex';
    document.getElementById('newClientName').focus();
}

function hideAddClientForm() {
    document.getElementById('addClientBtn').style.display = 'inline-block';
    document.getElementById('newClientForm').style.display = 'none';
    document.getElementById('newClientName').value = '';
}

async function saveNewClient() {
    const clientName = document.getElementById('newClientName').value.trim();
    
    if (!clientName) {
        showMessage('Please enter a client name', 'error');
        return;
    }
    
    if (clients.includes(clientName)) {
        showMessage('Client already exists', 'error');
        return;
    }
    
    try {
        showLoading();
        const placeholderAccount = {
            client: clientName,
            email: 'example@email.com',
            date: getTodayFormatted(),
            orderNumbers: ['', '', '', '', '']
        };
        
        const id = await saveAccount(placeholderAccount);
        accounts.push({ id, ...placeholderAccount });
        clients.push(clientName);
        
        hideLoading();
        hideAddClientForm();
        showMessage(`Client "${clientName}" added successfully`);
        renderAccounts();
        
        setTimeout(() => {
            editAccount(id);
        }, 100);
    } catch (error) {
        hideLoading();
        showMessage('Error adding client', 'error');
    }
}

// Modal Functions
function showConfirmModal(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    const messageEl = document.getElementById('confirmMessage');
    
    messageEl.textContent = message;
    modal.style.display = 'block';
    
    document.getElementById('confirmYes').onclick = () => {
        modal.style.display = 'none';
        onConfirm();
    };
    
    document.getElementById('confirmNo').onclick = () => {
        modal.style.display = 'none';
    };
}

// Debounce function for auto-search
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Event Listeners
document.addEventListener('DOMContentLoaded', function() {
    loadAccounts();
    
    // Add event delegation for order copy buttons
    document.addEventListener('click', handleOrderCopy);
    
    document.getElementById('addClientBtn').addEventListener('click', showAddClientForm);
    document.getElementById('saveClientBtn').addEventListener('click', saveNewClient);
    document.getElementById('cancelClientBtn').addEventListener('click', hideAddClientForm);
    
    document.getElementById('newClientName').addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            saveNewClient();
        }
    });
    
    document.getElementById('searchBtn').addEventListener('click', performSearch);
    document.getElementById('clearSearchBtn').addEventListener('click', clearSearch);
    
    // Order search functionality
    document.getElementById('orderSearchBtn').addEventListener('click', searchByOrderNumber);
    document.getElementById('clearOrderSearchBtn').addEventListener('click', clearOrderSearch);
    
    document.getElementById('orderSearchInput').addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            searchByOrderNumber();
        }
    });
    
    // Auto-search on input
    ['emailSearchInput', 'dateSearchInput', 'clientSearchInput'].forEach(inputId => {
        document.getElementById(inputId).addEventListener('input', debounce(performSearch, 300));
    });
    
    document.getElementById('orderSearchInput').addEventListener('input', debounce(searchByOrderNumber, 300));
    
    ['emailSearchInput', 'dateSearchInput', 'clientSearchInput'].forEach(inputId => {
        document.getElementById(inputId).addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                performSearch();
            }
        });
    });
    
    document.getElementById('exportCSVBtn').addEventListener('click', exportToCSV);
    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
    document.getElementById('exportJSONBtn').addEventListener('click', exportToJSON);
    
    document.getElementById('processBulkUpload').addEventListener('click', processBulkUpload);
    document.getElementById('cancelBulkUpload').addEventListener('click', closeBulkUploadModal);
    
    document.querySelectorAll('.close').forEach(closeBtn => {
        closeBtn.addEventListener('click', function() {
            this.closest('.modal').style.display = 'none';
        });
    });
    
    window.addEventListener('click', function(event) {
        const modals = document.querySelectorAll('.modal');
        modals.forEach(modal => {
            if (event.target === modal) {
                modal.style.display = 'none';
            }
        });
    });
    
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal').forEach(modal => {
                modal.style.display = 'none';
            });
        }
        
        if (e.ctrlKey && e.key === 'n') {
            e.preventDefault();
            showAddClientForm();
        }
        
        if (e.ctrlKey && e.key === 'f') {
            e.preventDefault();
            document.getElementById('emailSearchInput').focus();
        }
        
        if (e.ctrlKey && e.key === 'e') {
            e.preventDefault();
            exportToCSV();
        }
    });
});

// Error Handling
window.addEventListener('error', function(e) {
    console.error('Global error:', e.error);
    showMessage('An unexpected error occurred', 'error');
});