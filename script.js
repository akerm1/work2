/**************************************************************
 * script.js - MODIFIED VERSION
 *
 * Changes implemented:
 * 1. Improved edit UI for order inputs (larger, vertical layout)
 * 2. Auto-decrement order durations monthly
 * 3. Fixed bulk search for comma-separated emails
 * 4. Sort accounts with missing orders to top
 *
 ****************************************************************/
/* ============================================================
   SECTION: Firebase Configuration & Initialization
   ============================================================ */
const firebaseConfig = {
    apiKey: "AIzaSyDznYcUtQWRD7QqYBDr1QupUMfVqZnfGEE",
    authDomain: "my-work-82778.firebaseapp.com",
    projectId: "my-work-82778",
    storageBucket: "my-work-82778.appspot.com",
    messagingSenderId: "1070444118182",
    appId: "1:1070444118182:web:bae373255bd124d3a2b467"
};
let db = null;
try {
    firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();
    console.log('✅ Firebase initialized');
} catch (err) {
    console.error('❌ Firebase init error:', err);
    showError('Failed to connect to Firebase. Check your configuration in script.js.');
}

/* ============================================================
   SECTION: Global State
   ============================================================ */
const MAX_ORDER_SLOTS = 5;
let accounts = [];
let clients = [];
let filteredAccounts = [];
let isSearchActive = false;
let expandedClients = new Set();
let currentBulkClient = null;
let moveEmailTargetAccountId = null;

/* ============================================================
   SECTION: Utility Helpers
   ============================================================ */
function formatDateForDisplay(dateString) {
    if (!dateString) return '';
    const parts = String(dateString).split('-');
    if (parts.length === 3) {
        return `${parts[2].padStart(2,'0')}/${parts[1].padStart(2,'0')}`;
    }
    const d = new Date(dateString);
    if (d instanceof Date && !isNaN(d)) {
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth()+1).padStart(2, '0');
        return `${day}/${month}`;
    }
    return dateString;
}

function formatDateForStorage(dateString) {
    if (!dateString) return getTodayFormatted();
    return dateString;
}

function getTodayFormatted() {
    const t = new Date();
    const y = t.getFullYear();
    const m = String(t.getMonth()+1).padStart(2,'0');
    const d = String(t.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
}

function getTomorrowFormatted() {
    const t = new Date();
    t.setDate(t.getDate() + 1);
    const y = t.getFullYear();
    const m = String(t.getMonth()+1).padStart(2,'0');
    const d = String(t.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
}

function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return String(text)
        .replace(/&/g,'&amp;')
        .replace(/</g,'<')
        .replace(/>/g,'>')
        .replace(/"/g,'&quot;')
        .replace(/'/g,'&#039;');
}

function sanitizeOrderNumber(order) {
    return String(order || '').replace(/[^a-zA-Z0-9\s\-_@.]/g, '').trim().substring(0, 100);
}

/* ============================================================
   SECTION: UI Helpers
   ============================================================ */
function showMessage(message, type = 'success') {
    const messageEl = document.createElement('div');
    messageEl.className = `message ${type}`;
    messageEl.textContent = message;
    document.body.appendChild(messageEl);
    setTimeout(() => {
        try { messageEl.remove(); } catch (e) {}
    }, 4000);
}

function showError(message) {
    let errorBox = document.getElementById('errorBox');
    if (!errorBox) {
        errorBox = document.createElement('div');
        errorBox.id = 'errorBox';
        errorBox.style.position = 'fixed';
        errorBox.style.top = '10px';
        errorBox.style.left = '50%';
        errorBox.style.transform = 'translateX(-50%)';
        errorBox.style.background = '#ff4d4f';
        errorBox.style.color = '#fff';
        errorBox.style.padding = '10px 14px';
        errorBox.style.borderRadius = '6px';
        errorBox.style.zIndex = '99999';
        document.body.appendChild(errorBox);
    }
    errorBox.textContent = message;
}

function showLoading() {
    const el = document.getElementById('loadingSpinner');
    if (el) el.style.display = 'flex';
}

function hideLoading() {
    const el = document.getElementById('loadingSpinner');
    if (el) el.style.display = 'none';
}

function copyToClipboardSafe(text, successMessage='Copied to clipboard') {
    if (!text) {
        showMessage('Nothing to copy', 'error');
        return;
    }
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => showMessage(successMessage)).catch(() => fallbackCopy(text, successMessage));
    } else {
        fallbackCopy(text, successMessage);
    }
}

function fallbackCopy(text, successMessage) {
    try {
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.left = '-9999px';
        document.body.appendChild(ta);
        ta.focus(); ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showMessage(successMessage);
    } catch (err) {
        console.error('copy error', err);
        showMessage('Copy failed', 'error');
    }
}

/* ============================================================
   SECTION: Order rendering
   ============================================================ */
function renderOrderNumbers(orderNumbers, orderDurations) {
    orderNumbers = Array.isArray(orderNumbers) ? orderNumbers.slice(0, MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('');
    orderDurations = Array.isArray(orderDurations) ? orderDurations.slice(0, MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('1');
    while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
    while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');
    
    // Auto-decrement durations based on expiration date
    const accountDate = arguments[2]; // Pass expiration date as third argument
    if (accountDate) {
        const today = new Date();
        const expDate = new Date(accountDate);
        const monthsDiff = (expDate.getFullYear() - today.getFullYear()) * 12 + (expDate.getMonth() - today.getMonth());
        
        // Only decrement if expiration is in the past or current month
        if (monthsDiff <= 0) {
            orderDurations = orderDurations.map(dur => {
                const numDur = parseInt(dur) || 1;
                return Math.max(1, numDur - Math.abs(monthsDiff)).toString();
            });
        }
    }
    
    const parts = orderNumbers.map((order, i) => {
        if (order && String(order).trim()) {
            return `<div class="order-number"><span class="order-text">${escapeHtml(order)}</span> <span class="order-duration">(${escapeHtml(orderDurations[i])}M)</span> <button class="order-copy-btn" data-copy="${escapeHtml(order)}" title="Copy order">📋</button></div>`;
        } else {
            return `<div class="order-number empty"><span class="order-text">Order ${i+1}</span></div>`;
        }
    });
    return `<div class="order-numbers">${parts.join('')}</div>`;
}

function renderOrderInputs(orderNumbers, orderDurations, accountId) {
    orderNumbers = Array.isArray(orderNumbers) ? orderNumbers.slice(0,MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('');
    orderDurations = Array.isArray(orderDurations) ? orderDurations.slice(0,MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('1');
    while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
    while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');
    const parts = orderNumbers.map((order, i) => {
        const dur = orderDurations[i] || '1';
        return `
            <div class="order-with-duration">
                <input type="text" class="order-input" value="${escapeHtml(order)}" placeholder="Order ${i+1}" data-account="${escapeHtml(accountId)}" data-index="${i}" maxlength="100">
                <select class="order-duration" data-account="${escapeHtml(accountId)}" data-index="${i}">
                    <option value="1" ${dur==='1'?'selected':''}>1M</option>
                    <option value="2" ${dur==='2'?'selected':''}>2M</option>
                    <option value="3" ${dur==='3'?'selected':''}>3M</option>
                </select>
            </div>
        `;
    });
    return `<div class="order-inputs">${parts.join('')}</div>`;
}

/* ============================================================
   SECTION: Firestore CRUD Helpers
   ============================================================ */
async function loadAccounts() {
    if (!db) {
        showError('Firestore not initialized.');
        return;
    }
    try {
        showLoading();
        const snapshot = await db.collection('accounts').get();
        accounts = [];
        snapshot.forEach(doc => {
            const data = doc.data() || {};
            const orderNumbers = Array.isArray(data.orderNumbers) ? data.orderNumbers.slice(0, MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('');
            const orderDurations = Array.isArray(data.orderDurations) ? data.orderDurations.slice(0, MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('1');
            while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
            while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');
            const account = {
                id: doc.id,
                client: data.client || 'Unknown',
                email: data.email || '',
                date: data.date || getTodayFormatted(),
                orderNumbers: orderNumbers,
                orderDurations: orderDurations
            };
            accounts.push(account);
        });
        clients = [...new Set(accounts.map(a => a.client))].filter(Boolean);
        renderAccounts();
        renderExpiringAccounts();
        hideLoading();
        console.log('Loaded accounts:', accounts.length);
    } catch (err) {
        hideLoading();
        console.error('loadAccounts error', err);
        showMessage('Error loading accounts', 'error');
    }
}

async function saveAccount(accountData) {
    if (!db) throw new Error('Firestore not available');
    try {
        const docRef = await db.collection('accounts').add({
            client: accountData.client,
            email: accountData.email,
            date: formatDateForStorage(accountData.date),
            orderNumbers: accountData.orderNumbers || Array(MAX_ORDER_SLOTS).fill(''),
            orderDurations: accountData.orderDurations || Array(MAX_ORDER_SLOTS).fill('1')
        });
        return docRef.id;
    } catch (err) {
        console.error('saveAccount error', err);
        throw err;
    }
}

async function updateAccount(id, accountData) {
    if (!db) throw new Error('Firestore not available');
    try {
        await db.collection('accounts').doc(id).update({
            client: accountData.client,
            email: accountData.email,
            date: formatDateForStorage(accountData.date),
            orderNumbers: accountData.orderNumbers || Array(MAX_ORDER_SLOTS).fill(''),
            orderDurations: accountData.orderDurations || Array(MAX_ORDER_SLOTS).fill('1')
        });
    } catch (err) {
        console.error('updateAccount error', err);
        throw err;
    }
}

async function deleteAccount(id) {
    if (!db) throw new Error('Firestore not available');
    try {
        await db.collection('accounts').doc(id).delete();
    } catch (err) {
        console.error('deleteAccount error', err);
        throw err;
    }
}

async function bulkSaveAccounts(accountsData) {
    if (!db) throw new Error('Firestore not available');
    try {
        const batch = db.batch();
        accountsData.forEach(acc => {
            const docRef = db.collection('accounts').doc();
            batch.set(docRef, {
                client: acc.client,
                email: acc.email,
                date: formatDateForStorage(acc.date),
                orderNumbers: acc.orderNumbers || Array(MAX_ORDER_SLOTS).fill(''),
                orderDurations: acc.orderDurations || Array(MAX_ORDER_SLOTS).fill('1')
            });
        });
        await batch.commit();
    } catch (err) {
        console.error('bulkSaveAccounts error', err);
        throw err;
    }
}

/* ============================================================
   SECTION: Rendering Functions
   ============================================================ */
function renderAccounts() {
    const container = document.getElementById('clientsContainer');
    if (!container) return;
    container.innerHTML = '';
    
    // Group accounts by client
    const groups = {};
    accounts.forEach(acc => {
        if (!groups[acc.client]) groups[acc.client] = [];
        groups[acc.client].push(acc);
    });
    
    // Sort accounts within each client: missing orders first
    Object.keys(groups).forEach(client => {
        groups[client].sort((a, b) => {
            const aHasOrders = a.orderNumbers.some(o => o.trim());
            const bHasOrders = b.orderNumbers.some(o => o.trim());
            if (!aHasOrders && bHasOrders) return -1;
            if (aHasOrders && !bHasOrders) return 1;
            return 0;
        });
    });
    
    const clientNames = Object.keys(groups).sort((a,b) => a.localeCompare(b));
    if (clientNames.length === 0) {
        container.innerHTML = '<div class="no-data">No accounts found. Add your first client!</div>';
        return;
    }
    
    clientNames.forEach(clientName => {
        const clientAccounts = groups[clientName];
        const section = document.createElement('section');
        section.className = 'client-section';
        const isExpanded = expandedClients.has(clientName);
        const displayAccounts = isExpanded ? clientAccounts : clientAccounts.slice(0, 5);
        const showToggle = clientAccounts.length > 5;
        
        // Header
        const header = document.createElement('div');
        header.className = 'client-header';
        header.innerHTML = `
            <div class="client-name">${escapeHtml(clientName)} (${clientAccounts.length} account${clientAccounts.length !== 1 ? 's' : ''})</div>
            <div class="client-actions">
                <button class="btn" onclick="addNewAccount('${escapeForJs(clientName)}')">+ Add New Account</button>
                <button class="btn" onclick="openBulkUpload('${escapeForJs(clientName)}')">+ Bulk Upload Emails</button>
                <button class="btn" onclick="copyToClipboardSafe('${clientAccounts.map(a => a.email).join(', ')}','Copied all emails')">Copy All Emails</button>
            </div>
        `;
        
        // Table
        const table = document.createElement('div');
        table.className = 'table-container';
        const innerTable = document.createElement('table');
        innerTable.innerHTML = `
            <thead>
                <tr>
                    <th>Email</th>
                    <th>Expiration Date</th>
                    <th>Order Numbers</th>
                    <th>Actions</th>
                </tr>
            </thead>
            <tbody id="client-${escapeHtml(clientName).replace(/\s+/g,'-')}">
            </tbody>
        `;
        table.appendChild(innerTable);
        section.appendChild(header);
        section.appendChild(table);
        
        // Append rows
        const tbody = innerTable.querySelector('tbody');
        displayAccounts.forEach(account => {
            const tr = document.createElement('tr');
            tr.className = 'account-row';
            tr.dataset.id = account.id;
            tr.innerHTML = `
                <td class="email-cell">${escapeHtml(account.email)}</td>
                <td class="date-cell">${formatDateForDisplay(account.date)}</td>
                <td class="orders-cell">${renderOrderNumbers(account.orderNumbers, account.orderDurations, account.date)}</td>
                <td class="actions-cell">
                    <div class="action-buttons">
                        <button class="btn" onclick="editAccount('${escapeForJs(account.id)}')">Edit</button>
                        <button class="btn btn-danger" onclick="confirmDeleteAccount('${escapeForJs(account.id)}')">Delete</button>
                        <button class="btn" onclick="copyToClipboardSafe('${escapeHtml(account.email)}')">Copy</button>
                        <button class="btn" onclick="openMoveEmailModal('${escapeForJs(account.id)}')">Move Email</button>
                    </div>
                </td>
            `;
            tbody.appendChild(tr);
        });
        
        // Show toggle
        if (showToggle) {
            const toggleDiv = document.createElement('div');
            toggleDiv.className = 'toggle-container';
            toggleDiv.innerHTML = `<button class="btn" onclick="toggleClientAccounts('${escapeForJs(clientName)}')">${isExpanded ? 'Show Less' : 'Show More'}</button>`;
            section.appendChild(toggleDiv);
        }
        container.appendChild(section);
    });
}

function toggleClientAccounts(clientName) {
    if (expandedClients.has(clientName)) expandedClients.delete(clientName);
    else expandedClients.add(clientName);
    renderAccounts();
}

function renderExpiringAccounts() {
    const tomorrow = getTomorrowFormatted();
    const expiring = accounts.filter(acc => acc.date === tomorrow);
    const tbody = document.getElementById('expiringTableBody');
    const noDiv = document.getElementById('noExpiringAccounts');
    if (!tbody) return;
    if (expiring.length === 0) {
        tbody.innerHTML = '';
        if (noDiv) noDiv.style.display = 'block';
    } else {
        if (noDiv) noDiv.style.display = 'none';
        tbody.innerHTML = expiring.map(acc => `
            <tr class="account-row" data-id="${escapeHtml(acc.id)}">
                <td>${escapeHtml(acc.client)}</td>
                <td>${escapeHtml(acc.email)}</td>
                <td>${formatDateForDisplay(acc.date)}</td>
                <td>${renderOrderNumbers(acc.orderNumbers, acc.orderDurations, acc.date)}</td>
                <td>
                    <div class="action-buttons">
                        <button class="btn" onclick="editAccount('${escapeForJs(acc.id)}')">Edit</button>
                        <button class="btn btn-danger" onclick="confirmDeleteAccount('${escapeForJs(acc.id)}')">Delete</button>
                        <button class="btn" onclick="copyToClipboardSafe('${escapeHtml(acc.email)}')">Copy</button>
                        <button class="btn" onclick="openMoveEmailModal('${escapeForJs(acc.id)}')">Move Email</button>
                    </div>
                </td>
            </tr>
        `).join('');
    }
}

function renderSearchResults() {
    const container = document.getElementById('searchResults');
    const body = document.getElementById('searchResultsBody');
    const countEl = document.getElementById('resultCount');
    if (!container || !body || !countEl) return;
    if (!filteredAccounts || filteredAccounts.length === 0) {
        body.innerHTML = '<tr><td colspan="5" class="no-data">No accounts found matching your search.</td></tr>';
        countEl.textContent = '0 results';
        container.style.display = 'block';
        return;
    }
    
    // Group filtered by client
    const groups = {};
    filteredAccounts.forEach(acc => {
        if (!groups[acc.client]) groups[acc.client] = [];
        groups[acc.client].push(acc);
    });
    
    // Sort accounts within each client: missing orders first
    Object.keys(groups).forEach(client => {
        groups[client].sort((a, b) => {
            const aHasOrders = a.orderNumbers.some(o => o.trim());
            const bHasOrders = b.orderNumbers.some(o => o.trim());
            if (!aHasOrders && bHasOrders) return -1;
            if (aHasOrders && !bHasOrders) return 1;
            return 0;
        });
    });
    
    const htmlParts = [];
    Object.keys(groups).sort().forEach(client => {
        htmlParts.push(`
            <tr><td colspan="5" class="client-header">
                <div class="client-name">${escapeHtml(client)} (${groups[client].length})</div>
                <div><button class="btn" onclick="copyToClipboardSafe('${groups[client].map(a=>a.email).join(', ')}','Copied client emails')">Copy All</button></div>
            </td></tr>
        `);
        groups[client].forEach(account => {
            htmlParts.push(`
                <tr class="account-row" data-id="${escapeHtml(account.id)}">
                    <td>${escapeHtml(account.client)}</td>
                    <td>${escapeHtml(account.email)}</td>
                    <td>${formatDateForDisplay(account.date)}</td>
                    <td>${renderOrderNumbers(account.orderNumbers, account.orderDurations, account.date)}</td>
                    <td>
                        <div class="action-buttons">
                            <button class="btn" onclick="editAccount('${escapeForJs(account.id)}')">Edit</button>
                            <button class="btn btn-danger" onclick="confirmDeleteAccount('${escapeForJs(account.id)}')">Delete</button>
                            <button class="btn" onclick="copyToClipboardSafe('${escapeHtml(account.email)}')">Copy</button>
                            <button class="btn" onclick="openMoveEmailModal('${escapeForJs(account.id)}')">Move Email</button>
                        </div>
                    </td>
                </tr>
            `);
        });
    });
    body.innerHTML = htmlParts.join('');
    countEl.textContent = `${filteredAccounts.length} result${filteredAccounts.length !== 1 ? 's' : ''} found`;
    container.style.display = 'block';
}

/* ============================================================
   SECTION: Account Management
   ============================================================ */
function addNewAccount(clientName) {
    const tbodyId = `client-${escapeHtml(clientName).replace(/\s+/g,'-')}`;
    const tbody = document.getElementById(tbodyId);
    if (!tbody) {
        (async () => {
            try {
                showLoading();
                const placeholder = {
                    client: clientName,
                    email: `example+${Date.now()}@example.com`,
                    date: getTodayFormatted(),
                    orderNumbers: Array(MAX_ORDER_SLOTS).fill(''),
                    orderDurations: Array(MAX_ORDER_SLOTS).fill('1')
                };
                const id = await saveAccount(placeholder);
                accounts.push({ id, ...placeholder });
                if (!clients.includes(clientName)) clients.push(clientName);
                hideLoading();
                renderAccounts();
                editAccount(id);
            } catch (err) {
                hideLoading();
                console.error('addNewAccount fallback error', err);
                showMessage('Error adding placeholder account', 'error');
            }
        })();
        return;
    }
    
    const newRow = document.createElement('tr');
    newRow.className = 'editing-row';
    newRow.innerHTML = `
        <td><input type="email" class="new-email" placeholder="Enter email"></td>
        <td><input type="date" class="new-date" value="${getTodayFormatted()}"></td>
        <td>${renderOrderInputs(Array(MAX_ORDER_SLOTS).fill(''), Array(MAX_ORDER_SLOTS).fill('1'), 'new')}</td>
        <td>
            <div class="action-buttons">
                <button class="btn" onclick="saveNewAccount('${escapeForJs(clientName)}', this)">Save</button>
                <button class="btn btn-secondary" onclick="cancelNewAccount(this)">Cancel</button>
            </div>
        </td>
    `;
    tbody.appendChild(newRow);
    const input = newRow.querySelector('.new-email');
    if (input) { input.focus(); input.select(); }
}

async function saveNewAccount(clientName, button) {
    const row = button.closest('tr');
    if (!row) return;
    const emailEl = row.querySelector('.new-email');
    const dateEl = row.querySelector('.new-date');
    const orderInputs = row.querySelectorAll('.order-input');
    const durationSelects = row.querySelectorAll('.order-duration');
    const email = emailEl ? emailEl.value.trim() : '';
    const date = dateEl ? dateEl.value : getTodayFormatted();
    if (!email) { showMessage('Please enter an email', 'error'); return; }
    if (!validateEmail(email)) { showMessage('Invalid email', 'error'); return; }
    if (!validateDate(date)) { showMessage('Invalid date', 'error'); return; }
    const orderNumbers = Array.from(orderInputs).map(inp => sanitizeOrderNumber(inp.value));
    while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
    const orderDurations = Array.from(durationSelects).map(sel => sel.value || '1');
    while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');
    const accountData = {
        client: clientName,
        email: email,
        date: date,
        orderNumbers: orderNumbers,
        orderDurations: orderDurations
    };
    try {
        showLoading();
        const id = await saveAccount(accountData);
        accounts.push({ id, ...accountData });
        if (!clients.includes(clientName)) clients.push(clientName);
        hideLoading();
        showMessage('Account added successfully');
        renderAccounts();
        if (isSearchActive) performSearch();
        renderExpiringAccounts();
    } catch (err) {
        hideLoading();
        console.error('saveNewAccount error', err);
        showMessage('Error saving account', 'error');
    }
}

function cancelNewAccount(button) {
    const row = button.closest('tr');
    if (row) row.remove();
}

function editAccount(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) {
        showMessage('Account not found', 'error');
        return;
    }
    
    let row = document.querySelector(`#searchResultsBody tr.account-row[data-id="${id}"]`);
    if (!row) row = document.querySelector(`tr.account-row[data-id="${id}"]`);
    if (!row) {
        renderAccounts();
        row = document.querySelector(`tr.account-row[data-id="${id}"]`);
        if (!row) {
            showMessage('Could not find account row to edit', 'error');
            return;
        }
    }
    
    row.dataset.originalEmail = row.querySelector('.email-cell') ? row.querySelector('.email-cell').innerText : (account.email || '');
    row.dataset.originalDate = row.querySelector('.date-cell') ? row.querySelector('.date-cell').innerText : formatDateForDisplay(account.date);
    row.dataset.originalOrders = row.querySelector('.orders-cell') ? row.querySelector('.orders-cell').innerHTML : renderOrderNumbers(account.orderNumbers, account.orderDurations, account.date);
    row.dataset.originalActions = row.querySelector('.actions-cell') ? row.querySelector('.actions-cell').innerHTML : '';
    
    const emailCell = row.querySelector('.email-cell') || row.children[0];
    const dateCell = row.querySelector('.date-cell') || row.children[1];
    const ordersCell = row.querySelector('.orders-cell') || row.children[2];
    const actionsCell = row.querySelector('.actions-cell') || row.children[3];
    if (!emailCell || !dateCell || !ordersCell || !actionsCell) {
        showMessage('Edit failed: table structure unexpected', 'error');
        return;
    }
    
    emailCell.innerHTML = `<input type="email" class="edit-email" value="${escapeHtml(account.email)}">`;
    dateCell.innerHTML = `<input type="date" class="edit-date" value="${account.date || getTodayFormatted()}">`;
    ordersCell.innerHTML = renderOrderInputs(account.orderNumbers, account.orderDurations, account.id);
    actionsCell.innerHTML = `<div class="action-buttons"><button class="btn" onclick="saveAccountEdit('${escapeForJs(account.id)}')">Save</button><button class="btn btn-secondary" onclick="cancelAccountEdit('${escapeForJs(account.id)}')">Cancel</button></div>`;
    row.classList.add('editing-row');
    const e = row.querySelector('.edit-email');
    if (e) { e.focus(); e.select(); }
}

async function saveAccountEdit(id) {
    let row = document.querySelector(`#searchResultsBody tr.account-row[data-id="${id}"]`);
    if (!row) row = document.querySelector(`tr.account-row[data-id="${id}"]`);
    if (!row) {
        showMessage('Could not find account row to save', 'error');
        return;
    }
    const emailEl = row.querySelector('.edit-email');
    const dateEl = row.querySelector('.edit-date');
    const orderInputs = row.querySelectorAll('.order-input');
    const durationSelects = row.querySelectorAll('.order-duration');
    if (!emailEl || !dateEl) { showMessage('Edit fields missing', 'error'); return; }
    const email = emailEl.value.trim();
    const date = dateEl.value;
    if (!validateEmail(email)) { showMessage('Invalid email', 'error'); return; }
    if (!validateDate(date)) { showMessage('Invalid date', 'error'); return; }
    const orderNumbers = Array.from(orderInputs).map(i => sanitizeOrderNumber(i.value));
    while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
    const orderDurations = Array.from(durationSelects).map(s => s.value || '1');
    while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');
    const account = accounts.find(a => a.id === id);
    if (!account) { showMessage('Account not found in memory', 'error'); return; }
    const updated = {
        client: account.client,
        email: email,
        date: date || getTodayFormatted(),
        orderNumbers: orderNumbers,
        orderDurations: orderDurations
    };
    try {
        showLoading();
        await updateAccount(id, updated);
        Object.assign(account, updated);
        hideLoading();
        showMessage('Account updated');
        if (isSearchActive) renderSearchResults(); else renderAccounts();
        renderExpiringAccounts();
    } catch (err) {
        hideLoading();
        console.error('saveAccountEdit error', err);
        showMessage('Error updating account', 'error');
    }
}

function cancelAccountEdit(id) {
    let row = document.querySelector(`#searchResultsBody tr.account-row[data-id="${id}"]`);
    if (!row) row = document.querySelector(`tr.account-row[data-id="${id}"]`);
    if (!row) {
        if (isSearchActive) renderSearchResults(); else renderAccounts();
        return;
    }
    const account = accounts.find(a => a.id === id);
    if (account) {
        const emailCell = row.querySelector('.email-cell') || row.children[0];
        const dateCell = row.querySelector('.date-cell') || row.children[1];
        const ordersCell = row.querySelector('.orders-cell') || row.children[2];
        const actionsCell = row.querySelector('.actions-cell') || row.children[3];
        if (emailCell) emailCell.innerHTML = escapeHtml(account.email);
        if (dateCell) dateCell.innerHTML = formatDateForDisplay(account.date);
        if (ordersCell) ordersCell.innerHTML = renderOrderNumbers(account.orderNumbers, account.orderDurations, account.date);
        if (actionsCell) actionsCell.innerHTML = `
            <div class="action-buttons">
                <button class="btn" onclick="editAccount('${escapeForJs(account.id)}')">Edit</button>
                <button class="btn btn-danger" onclick="confirmDeleteAccount('${escapeForJs(account.id)}')">Delete</button>
                <button class="btn" onclick="copyToClipboardSafe('${escapeHtml(account.email)}')">Copy</button>
                <button class="btn" onclick="openMoveEmailModal('${escapeForJs(account.id)}')">Move Email</button>
            </div>
        `;
    } else {
        if (isSearchActive) renderSearchResults(); else renderAccounts();
    }
    row.classList.remove('editing-row');
}

function confirmDeleteAccount(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) { showMessage('Account not found', 'error'); return; }
    showConfirmModal(`Delete account for ${account.email}?`, () => deleteAccountConfirmed(id));
}

async function deleteAccountConfirmed(id) {
    try {
        showLoading();
        await deleteAccount(id);
        accounts = accounts.filter(a => a.id !== id);
        hideLoading();
        showMessage('Account deleted');
        renderAccounts();
        if (isSearchActive) performSearch();
        renderExpiringAccounts();
    } catch (err) {
        hideLoading();
        console.error('deleteAccountConfirmed error', err);
        showMessage('Error deleting account', 'error');
    }
}

/* ============================================================
   SECTION: Bulk Upload Handlers
   ============================================================ */
function openBulkUpload(clientName) {
    currentBulkClient = clientName;
    const modal = document.getElementById('bulkUploadModal');
    const input = document.getElementById('bulkEmailInput');
    if (input) input.value = '';
    if (modal) modal.style.display = 'block';
}

function closeBulkUploadModal() {
    currentBulkClient = null;
    const modal = document.getElementById('bulkUploadModal');
    if (modal) modal.style.display = 'none';
}

async function processBulkUpload() {
    const input = document.getElementById('bulkEmailInput');
    if (!input) { showMessage('Bulk input not found', 'error'); return; }
    const raw = input.value.trim();
    if (!raw) { showMessage('Please paste or type emails', 'error'); return; }
    const lines = raw.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const accountsToAdd = [];
    const errors = [];
    lines.forEach((line, idx) => {
        const parts = line.split(',').map(p => p.trim());
        const email = parts[0] || '';
        let date = parts[1] || getTodayFormatted();
        const orderNumbers = [];
        const orderDurations = [];
        if (parts.length > 2) {
            for (let i = 2; i < parts.length && orderNumbers.length < MAX_ORDER_SLOTS; i++) {
                const piece = parts[i];
                if (!piece) {
                    orderNumbers.push('');
                    orderDurations.push('1');
                    continue;
                }
                const sub = piece.split('|').map(s => s.trim());
                const ord = sanitizeOrderNumber(sub[0] || '');
                const dur = (sub[1] && ['1','2','3'].includes(sub[1])) ? sub[1] : '1';
                orderNumbers.push(ord);
                orderDurations.push(dur);
            }
        }
        while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
        while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');
        if (!validateEmail(email)) {
            errors.push(`Line ${idx+1}: invalid email "${email}"`);
            return;
        }
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (date && !dateRegex.test(date)) {
            errors.push(`Line ${idx+1}: invalid date "${date}" (use YYYY-MM-DD)`);
            return;
        }
        accountsToAdd.push({
            client: currentBulkClient || 'Unknown',
            email: email,
            date: date || getTodayFormatted(),
            orderNumbers: orderNumbers,
            orderDurations: orderDurations
        });
    });
    if (errors.length) {
        showMessage(errors.join('\n'), 'error');
        return;
    }
    if (accountsToAdd.length === 0) {
        showMessage('No valid accounts to upload', 'error');
        return;
    }
    try {
        showLoading();
        await bulkSaveAccounts(accountsToAdd);
        await loadAccounts();
        hideLoading();
        closeBulkUploadModal();
        showMessage(`Uploaded ${accountsToAdd.length} accounts`);
    } catch (err) {
        hideLoading();
        console.error('processBulkUpload error', err);
        showMessage('Error uploading accounts', 'error');
    }
}

/* ============================================================
   SECTION: Move Email
   ============================================================ */
function openMoveEmailModal(accountId) {
    const account = accounts.find(a => a.id === accountId);
    if (!account) { showMessage('Account not found', 'error'); return; }
    moveEmailTargetAccountId = accountId;
    const modal = document.getElementById('moveEmailModal');
    const select = document.getElementById('moveClientSelect');
    if (!modal || !select) { showMessage('Move modal missing', 'error'); return; }
    select.innerHTML = '';
    const uniqueClients = Array.from(new Set(clients.concat(accounts.map(a=>a.client)))).filter(Boolean);
    if (!uniqueClients.includes(account.client)) uniqueClients.unshift(account.client);
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = '-- Select client --';
    select.appendChild(placeholder);
    uniqueClients.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        if (c === account.client) opt.selected = true;
        select.appendChild(opt);
    });
    const createOpt = document.createElement('option');
    createOpt.value = '__create_new__';
    createOpt.textContent = '➕ Create new client...';
    select.appendChild(createOpt);
    modal.style.display = 'block';
    select.onchange = function() {
        const val = select.value;
        let existingInline = document.getElementById('moveNewClientInline');
        if (val === '__create_new__') {
            if (!existingInline) {
                const input = document.createElement('input');
                input.id = 'moveNewClientInline';
                input.placeholder = 'Enter new client name';
                input.style.marginTop = '10px';
                input.style.padding = '6px';
                input.style.width = '100%';
                select.parentNode.insertBefore(input, select.nextSibling);
            }
        } else {
            if (existingInline) existingInline.remove();
        }
    };
}

async function confirmMoveEmailHandler() {
    const modal = document.getElementById('moveEmailModal');
    const select = document.getElementById('moveClientSelect');
    if (!select || !modal) { showMessage('Move modal missing', 'error'); return; }
    const val = select.value;
    let targetClient = val;
    if (val === '') {
        showMessage('Please select a client or choose Create new client', 'error');
        return;
    }
    if (val === '__create_new__') {
        const inline = document.getElementById('moveNewClientInline');
        if (!inline || !inline.value.trim()) {
            showMessage('Please enter new client name', 'error');
            return;
        }
        targetClient = inline.value.trim();
    }
    if (!moveEmailTargetAccountId) {
        showMessage('No account selected to move', 'error');
        return;
    }
    const account = accounts.find(a => a.id === moveEmailTargetAccountId);
    if (!account) { showMessage('Account not found', 'error'); return; }
    const oldClient = account.client;
    account.client = targetClient;
    try {
        showLoading();
        await updateAccount(account.id, account);
        if (!clients.includes(targetClient)) clients.push(targetClient);
        const stillOld = accounts.some(a => a.client === oldClient && a.id !== account.id);
        if (!stillOld) clients = clients.filter(c => c !== oldClient);
        hideLoading();
        showMessage(`Moved ${account.email} to ${targetClient}`);
        closeMoveEmailModal();
        renderAccounts();
        renderExpiringAccounts();
        if (isSearchActive) performSearch();
    } catch (err) {
        hideLoading();
        console.error('confirmMoveEmailHandler error', err);
        showMessage('Error moving email', 'error');
    }
}

function closeMoveEmailModal() {
    const modal = document.getElementById('moveEmailModal');
    if (!modal) return;
    const inline = document.getElementById('moveNewClientInline');
    if (inline) inline.remove();
    const select = document.getElementById('moveClientSelect');
    if (select) select.onchange = null;
    moveEmailTargetAccountId = null;
    modal.style.display = 'none';
}

/* ============================================================
   SECTION: Export Functions
   ============================================================ */
function exportToCSV() {
    const data = isSearchActive ? filteredAccounts : accounts;
    if (!data || data.length === 0) { showMessage('No accounts to export', 'error'); return; }
    const headers = ['Client','Email','Expiration','Order1','Dur1','Order2','Dur2','Order3','Dur3','Order4','Dur4','Order5','Dur5'];
    const rows = data.map(acc => {
        const cols = [acc.client, acc.email, formatDateForDisplay(acc.date)];
        for (let i=0;i<MAX_ORDER_SLOTS;i++) {
            cols.push(acc.orderNumbers && acc.orderNumbers[i] ? acc.orderNumbers[i] : '');
            cols.push(acc.orderDurations && acc.orderDurations[i] ? acc.orderDurations[i] : '1');
        }
        return cols.map(c => `"${String(c||'').replace(/"/g,'""')}"`).join(',');
    });
    const csv = [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `email-accounts-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showMessage('CSV exported');
}

function exportToExcel() {
    const data = isSearchActive ? filteredAccounts : accounts;
    if (!data || data.length === 0) { showMessage('No accounts to export', 'error'); return; }
    const headerCells = ['Client','Email','Expiration'];
    for (let i=0;i<MAX_ORDER_SLOTS;i++) {
        headerCells.push(`Order ${i+1}`, `Duration ${i+1}`);
    }
    const rowsXml = data.map(acc => {
        const cells = [
            `<Cell><Data ss:Type="String">${escapeHtml(acc.client)}</Data></Cell>`,
            `<Cell><Data ss:Type="String">${escapeHtml(acc.email)}</Data></Cell>`,
            `<Cell><Data ss:Type="String">${escapeHtml(formatDateForDisplay(acc.date))}</Data></Cell>`
        ];
        for (let i=0;i<MAX_ORDER_SLOTS;i++) {
            cells.push(`<Cell><Data ss:Type="String">${escapeHtml(acc.orderNumbers ? acc.orderNumbers[i] || '' : '')}</Data></Cell>`);
            cells.push(`<Cell><Data ss:Type="String">${escapeHtml(acc.orderDurations ? acc.orderDurations[i] || '1' : '1')}</Data></Cell>`);
        }
        return `<Row>${cells.join('')}</Row>`;
    }).join('\n');
    const xml = `<?xml version="1.0"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
      <Worksheet ss:Name="Accounts">
        <Table>
          <Row>${headerCells.map(h=>`<Cell><Data ss:Type="String">${escapeHtml(h)}</Data></Cell>`).join('')}</Row>
          ${rowsXml}
        </Table>
      </Worksheet>
    </Workbook>`;
    const blob = new Blob([xml], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `email-accounts-${new Date().toISOString().split('T')[0]}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showMessage('Excel exported');
}

function exportToJSON() {
    const data = isSearchActive ? filteredAccounts : accounts;
    if (!data || data.length === 0) { showMessage('No accounts to export', 'error'); return; }
    const payload = data.map(acc => ({
        client: acc.client,
        email: acc.email,
        expirationDate: formatDateForDisplay(acc.date),
        rawDate: acc.date,
        orderNumbers: acc.orderNumbers,
        orderDurations: acc.orderDurations
    }));
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `email-accounts-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showMessage('JSON exported');
}

/* ============================================================
   SECTION: Search (Fixed for comma-separated emails)
   ============================================================ */
function performSearch() {
    const input = document.getElementById('unifiedSearchInput');
    if (!input) { showMessage('Search input not found', 'error'); return; }
    let q = input.value.trim();
    if (!q) {
        clearSearch();
        return;
    }
    
    // Handle comma-separated emails
    const searchTerms = q.split(/[,;]/).map(s => s.trim().toLowerCase()).filter(s => s);
    filteredAccounts = accounts.filter(acc => {
        const email = (acc.email || '').toLowerCase();
        const client = (acc.client || '').toLowerCase();
        const orders = (acc.orderNumbers || []).map(o => (o||'').toLowerCase());
        
        return searchTerms.some(term => 
            email.includes(term) || 
            client.includes(term) || 
            orders.some(o => o.includes(term))
        );
    });
    
    isSearchActive = true;
    renderSearchResults();
    if (!filteredAccounts || filteredAccounts.length === 0) {
        showMessage('No accounts found', 'error');
    }
}

function clearSearch() {
    const input = document.getElementById('unifiedSearchInput');
    if (input) input.value = '';
    const results = document.getElementById('searchResults');
    if (results) results.style.display = 'none';
    isSearchActive = false;
    filteredAccounts = [];
}

/* ============================================================
   SECTION: Utilities
   ============================================================ */
function validateEmail(email) {
    if (!email) return false;
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
}

function validateDate(dateString) {
    if (!dateString) return true;
    const re = /^\d{4}-\d{2}-\d{2}$/;
    if (!re.test(dateString)) return false;
    const d = new Date(dateString);
    return d instanceof Date && !isNaN(d);
}

function escapeForJs(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

/* ============================================================
   SECTION: Event Delegation
   ============================================================ */
document.addEventListener('click', function(e) {
    const el = e.target;
    if (!el) return;
    if (el.classList && el.classList.contains('order-copy-btn')) {
        const v = el.getAttribute('data-copy') || '';
        if (v) copyToClipboardSafe(v, `Order "${v}" copied`);
        e.preventDefault();
        return;
    }
});

/* ============================================================
   SECTION: Confirm Modal
   ============================================================ */
function showConfirmModal(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    const messageEl = document.getElementById('confirmMessage');
    if (!modal || !messageEl) {
        if (confirm(message)) onConfirm();
        return;
    }
    messageEl.textContent = message;
    modal.style.display = 'block';
    document.getElementById('confirmYes').onclick = function() {
        modal.style.display = 'none';
        onConfirm();
    };
    document.getElementById('confirmNo').onclick = function() {
        modal.style.display = 'none';
    };
}

/* ============================================================
   SECTION: Event Wiring
   ============================================================ */
document.addEventListener('DOMContentLoaded', function() {
    if (db) loadAccounts();
    
    const unifiedSearchBtn = document.getElementById('unifiedSearchBtn');
    const unifiedSearchInput = document.getElementById('unifiedSearchInput');
    const clearUnifiedSearchBtn = document.getElementById('clearUnifiedSearchBtn');
    if (unifiedSearchBtn) unifiedSearchBtn.addEventListener('click', performSearch);
    if (clearUnifiedSearchBtn) clearUnifiedSearchBtn.addEventListener('click', clearSearch);
    if (unifiedSearchInput) {
        unifiedSearchInput.addEventListener('keypress', function(e) { if (e.key === 'Enter') performSearch(); });
        unifiedSearchInput.addEventListener('input', debounce(performSearch, 300));
    }
    
    const exportCSVBtn = document.getElementById('exportCSVBtn');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportJSONBtn = document.getElementById('exportJSONBtn');
    if (exportCSVBtn) exportCSVBtn.addEventListener('click', exportToCSV);
    if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
    if (exportJSONBtn) exportJSONBtn.addEventListener('click', exportToJSON);
    
    const processBulkBtn = document.getElementById('processBulkUpload');
    const cancelBulkBtn = document.getElementById('cancelBulkUpload');
    if (processBulkBtn) processBulkBtn.addEventListener('click', processBulkUpload);
    if (cancelBulkBtn) cancelBulkBtn.addEventListener('click', closeBulkUploadModal);
    
    document.querySelectorAll('.close').forEach(c => c.addEventListener('click', function() {
        const m = this.closest('.modal');
        if (m) m.style.display = 'none';
    }));
    
    const confirmMoveEmailBtn = document.getElementById('confirmMoveEmail');
    const cancelMoveEmailBtn = document.getElementById('cancelMoveEmail');
    if (confirmMoveEmailBtn) confirmMoveEmailBtn.addEventListener('click', confirmMoveEmailHandler);
    if (cancelMoveEmailBtn) cancelMoveEmailBtn.addEventListener('click', closeMoveEmailModal);
    
    window.addEventListener('click', function(e) {
        document.querySelectorAll('.modal').forEach(modal => {
            if (e.target === modal) modal.style.display = 'none';
        });
    });
    
    window.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal').forEach(m => m.style.display = 'none');
        }
        if (e.ctrlKey && e.key === 'f') {
            e.preventDefault();
            const input = document.getElementById('unifiedSearchInput');
            if (input) input.focus();
        }
        if (e.ctrlKey && e.key === 'e') {
            e.preventDefault();
            exportToCSV();
        }
    });
});

/* ============================================================
   SECTION: Debounce
   ============================================================ */
function debounce(fn, wait) {
    let t;
    return function(...args) {
        clearTimeout(t);
        t = setTimeout(() => fn.apply(this, args), wait);
    };
}

/* ============================================================
   SECTION: Global Error Handler
   ============================================================ */
window.addEventListener('error', function(e) {
    console.error('Global error:', e.error || e.message || e);
    showMessage('An unexpected error occurred', 'error');
});
