/**************************************************************
 * script.js
 *
 * Email Account Organizer - FULL script
 * - Uses Firestore (compat) for storage
 * - Unified search (email/client/order)
 * - Order durations (1/2/3 months)
 * - Move Email modal (in-site, suggests existing clients)
 * - Bulk upload modal
 * - Export (CSV / Excel / JSON)
 * - Edit / Add / Delete accounts
 * - Plenty of section headers and inline comments for clarity
 *
 * IMPORTANT:
 * - Replace the firebaseConfig below with your project's keys if needed.
 * - This code expects the following HTML IDs (from HTML you accepted):
 *   #unifiedSearchInput, #unifiedSearchBtn, #clearUnifiedSearchBtn,
 *   #clientsContainer, #expiringTableBody, #noExpiringAccounts,
 *   #exportCSVBtn, #exportExcelBtn, #exportJSONBtn,
 *   #bulkUploadModal, #bulkEmailInput, #processBulkUpload, #cancelBulkUpload,
 *   #confirmModal, #confirmMessage, #confirmYes, #confirmNo,
 *   #moveEmailModal, #moveClientSelect, #confirmMoveEmail, #cancelMoveEmail,
 *   #loadingSpinner, #searchResults, #searchResultsBody, #resultCount
 *
 ****************************************************************/

/* ============================================================
   SECTION: Firebase Configuration & Initialization
   ============================================================ */

// NOTE: keep your real API keys here; the example below uses placeholder values
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
    // initialize firebase compat
    firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();
    console.log('✅ Firebase initialized');
} catch (err) {
    console.error('❌ Firebase init error:', err);
    // show friendly error on page
    showError('Failed to connect to Firebase. Check your configuration in script.js.');
}

/* ============================================================
   SECTION: Global State
   ============================================================ */

const MAX_ORDER_SLOTS = 5; // keep 5 order slots like your original design

let accounts = [];           // all accounts loaded from Firestore
let clients = [];            // unique client names
let filteredAccounts = [];   // results when searching
let isSearchActive = false;  // whether search results are shown
let expandedClients = new Set(); // which clients are expanded in UI

let currentBulkClient = null;   // client for bulk upload modal
let moveEmailTargetAccountId = null; // account id currently being moved (for modal)

/* ============================================================
   SECTION: Utility Helpers (dates, sanitize, escape)
   ============================================================ */

/**
 * formatDateForDisplay()
 * - shows dd/mm/yyyy (safe and stable)
 */
function formatDateForDisplay(dateString) {
    if (!dateString) return '';
    const parts = String(dateString).split('-');
    if (parts.length === 3) {
        // assume yyyy-mm-dd
        return `${parts[2].padStart(2,'0')}/${parts[1].padStart(2,'0')}/${parts[0]}`;
    }
    // fallback to Date parsing
    const d = new Date(dateString);
    if (d instanceof Date && !isNaN(d)) {
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth()+1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
    }
    return dateString;
}

/**
 * formatDateForStorage()
 * - ensures YYYY-MM-DD or today
 */
function formatDateForStorage(dateString) {
    if (!dateString) return getTodayFormatted();
    return dateString;
}

/**
 * getTodayFormatted() / getTomorrowFormatted()
 */
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

/**
 * escapeHtml()
 * - simple html escape to prevent injection in rendered strings
 */
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return String(text)
        .replace(/&/g,'&amp;')
        .replace(/</g,'&lt;')
        .replace(/>/g,'&gt;')
        .replace(/"/g,'&quot;')
        .replace(/'/g,'&#039;');
}

/**
 * sanitizeOrderNumber()
 * - keep only useful characters
 */
function sanitizeOrderNumber(order) {
    return String(order || '').replace(/[^a-zA-Z0-9\s\-_@.]/g, '').trim().substring(0, 100);
}

/* ============================================================
   SECTION: UI Helpers - Notifications & Loading
   ============================================================ */

/* NOTE: This showMessage function restores the original style you had:
   - creates a `.message` element with class "success" or "error"
   - disappears after 4 seconds
*/
function showMessage(message, type = 'success') {
    const messageEl = document.createElement('div');
    messageEl.className = `message ${type}`;
    messageEl.textContent = message;
    document.body.appendChild(messageEl);

    setTimeout(() => {
        try { messageEl.remove(); } catch (e) {}
    }, 4000);
}

/* showError creates a visible fixed error box at top (used for critical failures) */
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

/* Loading spinner show/hide (assumes #loadingSpinner exists) */
function showLoading() {
    const el = document.getElementById('loadingSpinner');
    if (el) el.style.display = 'flex';
}
function hideLoading() {
    const el = document.getElementById('loadingSpinner');
    if (el) el.style.display = 'none';
}

/* Copy to clipboard safe method (clipboard API + fallback) */
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
   SECTION: Order rendering (numbers + durations)
   ============================================================ */

/**
 * renderOrderNumbers(orderNumbers, orderDurations)
 * - returns an HTML snippet showing orders with duration labels and copy buttons
 */
function renderOrderNumbers(orderNumbers, orderDurations) {
    orderNumbers = Array.isArray(orderNumbers) ? orderNumbers.slice(0, MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('');
    orderDurations = Array.isArray(orderDurations) ? orderDurations.slice(0, MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('1');

    // ensure length
    while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
    while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');

    const parts = orderNumbers.map((order, i) => {
        if (order && String(order).trim()) {
            return `<div class="order-number"><span class="order-text">${escapeHtml(order)}</span> <span class="order-duration">(${escapeHtml(orderDurations[i])}M)</span> <button class="order-copy-btn" data-copy="${escapeHtml(order)}" title="Copy order">📋</button></div>`;
        } else {
            return `<div class="order-number empty"><span class="order-text">Order ${i+1}</span></div>`;
        }
    });

    return `<div class="order-numbers">${parts.join('')}</div>`;
}

/**
 * renderOrderInputs(orderNumbers, orderDurations, accountId)
 * - renders input fields plus a select for duration (1/2/3)
 */
function renderOrderInputs(orderNumbers, orderDurations, accountId) {
    orderNumbers = Array.isArray(orderNumbers) ? orderNumbers.slice(0,MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('');
    orderDurations = Array.isArray(orderDurations) ? orderDurations.slice(0,MAX_ORDER_SLOTS) : Array(MAX_ORDER_SLOTS).fill('1');

    while (orderNumbers.length < MAX_ORDER_SLOTS) orderNumbers.push('');
    while (orderDurations.length < MAX_ORDER_SLOTS) orderDurations.push('1');

    const parts = orderNumbers.map((order, i) => {
        const dur = orderDurations[i] || '1';
        return `
            <div class="order-with-duration" style="display:flex;gap:6px;align-items:center;margin-bottom:6px;">
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

/**
 * loadAccounts()
 * - loads all accounts from firestore into 'accounts' array
 * - normalizes missing orderDurations for backwards compatibility
 */
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

/**
 * saveAccount(accountData)
 * - adds a new document
 * - returns doc id
 */
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

/**
 * updateAccount(id, accountData)
 * - updates the document with given id
 */
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

/**
 * deleteAccount(id)
 */
async function deleteAccount(id) {
    if (!db) throw new Error('Firestore not available');
    try {
        await db.collection('accounts').doc(id).delete();
    } catch (err) {
        console.error('deleteAccount error', err);
        throw err;
    }
}

/**
 * bulkSaveAccounts(accountsData)
 * - writes multiple docs in a batch
 */
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
   SECTION: Rendering - Accounts, Clients, Expiring & Search
   ============================================================ */

/**
 * renderAccounts()
 * - builds the clients container with tables grouped by client
 */
function renderAccounts() {
    const container = document.getElementById('clientsContainer');
    if (!container) return;
    container.innerHTML = '';

    // group accounts by client
    const groups = {};
    accounts.forEach(acc => {
        if (!groups[acc.client]) groups[acc.client] = [];
        groups[acc.client].push(acc);
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

        // header
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

        // table
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

        // append rows
        const tbody = innerTable.querySelector('tbody');
        displayAccounts.forEach(account => {
            const tr = document.createElement('tr');
            tr.className = 'account-row';
            tr.dataset.id = account.id;
            tr.innerHTML = `
                <td class="email-cell">${escapeHtml(account.email)}</td>
                <td class="date-cell">${formatDateForDisplay(account.date)}</td>
                <td class="orders-cell">${renderOrderNumbers(account.orderNumbers, account.orderDurations)}</td>
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

        // show toggle
        if (showToggle) {
            const toggleDiv = document.createElement('div');
            toggleDiv.className = 'toggle-container';
            toggleDiv.innerHTML = `<button class="btn" onclick="toggleClientAccounts('${escapeForJs(clientName)}')">${isExpanded ? 'Show Less' : 'Show More'}</button>`;
            section.appendChild(toggleDiv);
        }

        container.appendChild(section);
    });
}

/**
 * toggleClientAccounts(clientName)
 */
function toggleClientAccounts(clientName) {
    if (expandedClients.has(clientName)) expandedClients.delete(clientName);
    else expandedClients.add(clientName);
    renderAccounts();
}

/**
 * renderExpiringAccounts()
 * - shows accounts whose date is tomorrow
 */
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
                <td>${renderOrderNumbers(acc.orderNumbers, acc.orderDurations)}</td>
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

/**
 * renderSearchResults()
 * - shows filteredAccounts grouped by client
 */
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

    // group filtered by client
    const groups = {};
    filteredAccounts.forEach(acc => {
        if (!groups[acc.client]) groups[acc.client] = [];
        groups[acc.client].push(acc);
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
                    <td>${renderOrderNumbers(account.orderNumbers, account.orderDurations)}</td>
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
   SECTION: Account Add / Edit / Save / Cancel / Delete
   ============================================================ */

/**
 * addNewAccount(clientName)
 * - inserts a new editable row for adding an account under the client
 */
function addNewAccount(clientName) {
    const tbodyId = `client-${escapeHtml(clientName).replace(/\s+/g,'-')}`;
    const tbody = document.getElementById(tbodyId);
    // If tbody not present (client group not visible), just call renderAccounts then open edit for newly created placeholder
    if (!tbody) {
        // create placeholder account then edit
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

    // create a new row in the client tbody for immediate input
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

/**
 * saveNewAccount(clientName, button)
 */
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

/**
 * cancelNewAccount(button)
 */
function cancelNewAccount(button) {
    const row = button.closest('tr');
    if (row) row.remove();
}

/**
 * editAccount(id)
 * - converts an existing row to editable form
 */
function editAccount(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) {
        showMessage('Account not found', 'error');
        return;
    }

    // find the row in search results or main view
    let row = document.querySelector(`#searchResultsBody tr.account-row[data-id="${id}"]`);
    if (!row) row = document.querySelector(`tr.account-row[data-id="${id}"]`);
    if (!row) {
        // fallback: re-render and then find again
        renderAccounts();
        row = document.querySelector(`tr.account-row[data-id="${id}"]`);
        if (!row) {
            showMessage('Could not find account row to edit', 'error');
            return;
        }
    }

    // save original in data attributes for cancel
    row.dataset.originalEmail = row.querySelector('.email-cell') ? row.querySelector('.email-cell').innerText : (account.email || '');
    row.dataset.originalDate = row.querySelector('.date-cell') ? row.querySelector('.date-cell').innerText : formatDateForDisplay(account.date);
    row.dataset.originalOrders = row.querySelector('.orders-cell') ? row.querySelector('.orders-cell').innerHTML : renderOrderNumbers(account.orderNumbers, account.orderDurations);
    row.dataset.originalActions = row.querySelector('.actions-cell') ? row.querySelector('.actions-cell').innerHTML : '';

    // build edit UI
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

/**
 * saveAccountEdit(id)
 */
async function saveAccountEdit(id) {
    // find the row
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
        // update local model
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

/**
 * cancelAccountEdit(id)
 */
function cancelAccountEdit(id) {
    // try to find row
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
        if (ordersCell) ordersCell.innerHTML = renderOrderNumbers(account.orderNumbers, account.orderDurations);
        if (actionsCell) actionsCell.innerHTML = `
            <div class="action-buttons">
                <button class="btn" onclick="editAccount('${escapeForJs(account.id)}')">Edit</button>
                <button class="btn btn-danger" onclick="confirmDeleteAccount('${escapeForJs(account.id)}')">Delete</button>
                <button class="btn" onclick="copyToClipboardSafe('${escapeHtml(account.email)}')">Copy</button>
                <button class="btn" onclick="openMoveEmailModal('${escapeForJs(account.id)}')">Move Email</button>
            </div>
        `;
    } else {
        // if account not found, re-render overall
        if (isSearchActive) renderSearchResults(); else renderAccounts();
    }
    row.classList.remove('editing-row');
}

/**
 * confirmDeleteAccount(id)
 */
function confirmDeleteAccount(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) { showMessage('Account not found', 'error'); return; }
    showConfirmModal(`Delete account for ${account.email}?`, () => deleteAccountConfirmed(id));
}

/**
 * deleteAccountConfirmed(id)
 */
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

/**
 * openBulkUpload(clientName)
 */
function openBulkUpload(clientName) {
    currentBulkClient = clientName;
    const modal = document.getElementById('bulkUploadModal');
    const input = document.getElementById('bulkEmailInput');
    if (input) input.value = '';
    if (modal) modal.style.display = 'block';
}

/**
 * closeBulkUploadModal()
 */
function closeBulkUploadModal() {
    currentBulkClient = null;
    const modal = document.getElementById('bulkUploadModal');
    if (modal) modal.style.display = 'none';
}

/**
 * processBulkUpload()
 * Accepts lines:
 * email
 * email,YYYY-MM-DD
 * email,YYYY-MM-DD,order1,order2|2,order3|3,order4,order5
 * where orderX|Y sets duration Y for that order slot (1/2/3)
 */
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
                // allow order|duration
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
   SECTION: Move Email - Modal + Logic (in-site, suggests existing clients)
   ============================================================ */

/**
 * openMoveEmailModal(accountId)
 * - populates move modal with client list and selects current client by default
 */
function openMoveEmailModal(accountId) {
    const account = accounts.find(a => a.id === accountId);
    if (!account) { showMessage('Account not found', 'error'); return; }
    moveEmailTargetAccountId = accountId;

    const modal = document.getElementById('moveEmailModal');
    const select = document.getElementById('moveClientSelect');
    if (!modal || !select) { showMessage('Move modal missing', 'error'); return; }

    // clear select
    select.innerHTML = '';

    // option: keep current client as selected but also allow "Create new client" option
    // Build options from clients array (unique)
    const uniqueClients = Array.from(new Set(clients.concat(accounts.map(a=>a.client)))).filter(Boolean);
    // ensure current client is present
    if (!uniqueClients.includes(account.client)) uniqueClients.unshift(account.client);

    // add placeholder option
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = '-- Select client --';
    select.appendChild(placeholder);

    // add current client option first
    uniqueClients.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        if (c === account.client) opt.selected = true;
        select.appendChild(opt);
    });

    // add "Create new client" option
    const createOpt = document.createElement('option');
    createOpt.value = '__create_new__';
    createOpt.textContent = '➕ Create new client...';
    select.appendChild(createOpt);

    // show modal
    modal.style.display = 'block';

    // when create new selected, show a prompt inline (to keep everything inside site, we'll add an inline input)
    // We will create an inline input if needed
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

/**
 * confirmMoveEmailHandler()
 * - called when user clicks Confirm in move modal
 */
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

    // proceed to update the account's client
    const account = accounts.find(a => a.id === moveEmailTargetAccountId);
    if (!account) { showMessage('Account not found', 'error'); return; }

    const oldClient = account.client;
    account.client = targetClient;

    try {
        showLoading();
        await updateAccount(account.id, account);
        // update local clients list
        if (!clients.includes(targetClient)) clients.push(targetClient);
        // remove old client if no accounts left
        const stillOld = accounts.some(a => a.client === oldClient && a.id !== account.id);
        if (!stillOld) clients = clients.filter(c => c !== oldClient);
        hideLoading();
        showMessage(`Moved ${account.email} to ${targetClient}`);
        // close modal and cleanup
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

/**
 * closeMoveEmailModal()
 */
function closeMoveEmailModal() {
    const modal = document.getElementById('moveEmailModal');
    if (!modal) return;
    // cleanup inline input if exists
    const inline = document.getElementById('moveNewClientInline');
    if (inline) inline.remove();
    // clear select onchange handler
    const select = document.getElementById('moveClientSelect');
    if (select) select.onchange = null;
    moveEmailTargetAccountId = null;
    modal.style.display = 'none';
}

/* ============================================================
   SECTION: Export Functions (CSV / Excel / JSON)
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
        // escape each cell with double quotes and replace inner quotes
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
   SECTION: Search (Unified)
   ============================================================ */

/**
 * performSearch()
 * - uses #unifiedSearchInput and searches email, client, and order numbers
 */
function performSearch() {
    const input = document.getElementById('unifiedSearchInput');
    if (!input) { showMessage('Search input not found', 'error'); return; }
    const q = input.value.trim().toLowerCase();
    if (!q) {
        clearSearch();
        return;
    }

    filteredAccounts = accounts.filter(acc => {
        const email = (acc.email || '').toLowerCase();
        const client = (acc.client || '').toLowerCase();
        const orders = (acc.orderNumbers || []).map(o => (o||'').toLowerCase());
        return email.includes(q) || client.includes(q) || orders.some(o => o.includes(q));
    });

    isSearchActive = true;
    renderSearchResults();

    if (!filteredAccounts || filteredAccounts.length === 0) {
        showMessage('No accounts found', 'error');
    }
}

/**
 * clearSearch()
 */
function clearSearch() {
    const input = document.getElementById('unifiedSearchInput');
    if (input) input.value = '';
    const results = document.getElementById('searchResults');
    if (results) results.style.display = 'none';
    isSearchActive = false;
    filteredAccounts = [];
}

/* ============================================================
   SECTION: Small utilities (validation, escaping for onclick)
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
   SECTION: Order copy button handler (delegation)
   ============================================================ */
document.addEventListener('click', function(e) {
    const el = e.target;
    if (!el) return;
    // order copy button may be a button with class order-copy-btn
    if (el.classList && el.classList.contains('order-copy-btn')) {
        const v = el.getAttribute('data-copy') || '';
        if (v) copyToClipboardSafe(v, `Order "${v}" copied`);
        e.preventDefault();
        return;
    }
});

/* ============================================================
   SECTION: Confirm modal helper
   ============================================================ */

function showConfirmModal(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    const messageEl = document.getElementById('confirmMessage');
    if (!modal || !messageEl) {
        // fallback to built-in confirm
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
   SECTION: Event Wiring (DOMContentLoaded)
   ============================================================ */

document.addEventListener('DOMContentLoaded', function() {
    // initial load
    if (db) loadAccounts();

    // wire search
    const unifiedSearchBtn = document.getElementById('unifiedSearchBtn');
    const unifiedSearchInput = document.getElementById('unifiedSearchInput');
    const clearUnifiedSearchBtn = document.getElementById('clearUnifiedSearchBtn');
    if (unifiedSearchBtn) unifiedSearchBtn.addEventListener('click', performSearch);
    if (clearUnifiedSearchBtn) clearUnifiedSearchBtn.addEventListener('click', clearSearch);
    if (unifiedSearchInput) {
        unifiedSearchInput.addEventListener('keypress', function(e) { if (e.key === 'Enter') performSearch(); });
        // debounce input
        unifiedSearchInput.addEventListener('input', debounce(performSearch, 300));
    }

    // export buttons
    const exportCSVBtn = document.getElementById('exportCSVBtn');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportJSONBtn = document.getElementById('exportJSONBtn');
    if (exportCSVBtn) exportCSVBtn.addEventListener('click', exportToCSV);
    if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
    if (exportJSONBtn) exportJSONBtn.addEventListener('click', exportToJSON);

    // bulk upload buttons
    const processBulkBtn = document.getElementById('processBulkUpload');
    const cancelBulkBtn = document.getElementById('cancelBulkUpload');
    if (processBulkBtn) processBulkBtn.addEventListener('click', processBulkUpload);
    if (cancelBulkBtn) cancelBulkBtn.addEventListener('click', closeBulkUploadModal);

    // close icons
    document.querySelectorAll('.close').forEach(c => c.addEventListener('click', function() {
        const m = this.closest('.modal');
        if (m) m.style.display = 'none';
    }));

    // confirm modal click handlers are set when modal opened

    // move email modal confirm/cancel
    const confirmMoveEmailBtn = document.getElementById('confirmMoveEmail');
    const cancelMoveEmailBtn = document.getElementById('cancelMoveEmail');
    if (confirmMoveEmailBtn) confirmMoveEmailBtn.addEventListener('click', confirmMoveEmailHandler);
    if (cancelMoveEmailBtn) cancelMoveEmailBtn.addEventListener('click', closeMoveEmailModal);

    // click outside modal to close
    window.addEventListener('click', function(e) {
        document.querySelectorAll('.modal').forEach(modal => {
            if (e.target === modal) modal.style.display = 'none';
        });
    });

    // keyboard shortcuts
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
   SECTION: Debounce utility
   ============================================================ */
function debounce(fn, wait) {
    let t;
    return function(...args) {
        clearTimeout(t);
        t = setTimeout(() => fn.apply(this, args), wait);
    };
}

/* ============================================================
   SECTION: Legacy order search helpers (if your HTML still has orderSearchInput)
   ============================================================ */

function searchByOrderNumber() {
    const orderInput = document.getElementById('orderSearchInput');
    if (!orderInput) return;
    const q = orderInput.value.trim().toLowerCase();
    if (!q) {
        document.getElementById('orderResults') && (document.getElementById('orderResults').style.display = 'none');
        return;
    }
    const matches = accounts.filter(acc => acc.orderNumbers && acc.orderNumbers.some(o => (o||'').toLowerCase().includes(q)));
    renderOrderSearchResults(matches, q);
}
function renderOrderSearchResults(matches, query) {
    const resultsDiv = document.getElementById('orderResults');
    if (!resultsDiv) return;
    if (!matches || matches.length === 0) {
        resultsDiv.innerHTML = `<div class="no-data">No accounts found with order number containing "${escapeHtml(query)}"</div>`;
        resultsDiv.style.display = 'block';
        return;
    }
    const html = matches.map(acc => {
        const matching = acc.orderNumbers.filter(o => (o||'').toLowerCase().includes(query)).join(', ');
        return `
            <div class="order-match">
                <div class="order-match-header">${escapeHtml(acc.client)} - ${escapeHtml(acc.email)}</div>
                <div class="order-match-details">Matching Orders: <strong>${escapeHtml(matching)}</strong><br>All: ${escapeHtml((acc.orderNumbers||[]).filter(Boolean).join(', '))}<br>Exp: ${formatDateForDisplay(acc.date)}</div>
            </div>
        `;
    }).join('');
    resultsDiv.innerHTML = `<h4>Found ${matches.length} account${matches.length !== 1 ? 's' : ''}</h4>${html}`;
    resultsDiv.style.display = 'block';
}
function clearOrderSearch() {
    const orderInput = document.getElementById('orderSearchInput');
    if (orderInput) orderInput.value = '';
    const orderResults = document.getElementById('orderResults');
    if (orderResults) orderResults.style.display = 'none';
}

/* ============================================================
   SECTION: Safe global error handler
   ============================================================ */
window.addEventListener('error', function(e) {
    console.error('Global error:', e.error || e.message || e);
    showMessage('An unexpected error occurred', 'error');
});

/* ============================================================
   END OF FILE - script.js
   ============================================================ */
