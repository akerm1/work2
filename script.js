// Firebase Configuration
const firebaseConfig = {
    apiKey: "AIzaSyDznYcUtQWRD7QqYBDr1QupUMfVqZnfGEE",
    authDomain: "my-work-82778.firebaseapp.com",
    projectId: "my-work-82778",
    storageBucket: "my-work-82778.appspot.com",
    messagingSenderId: "1070444118182",
    appId: "1:1070444118182:web:bae373255bd124d3a2b467"
};

let db = null;

let accounts = [];
let clients = [];
let filteredAccounts = [];
let isSearchActive = false;
let expandedClients = new Set();
let currentReplacementAccountId = null;
let currentProblemAccountId = null;
let currentBulkClient = null;
let notificationInterval = null;
let notifiedAccountIds = new Set();
(function loadNotifiedToday() {
    try {
        const saved = JSON.parse(localStorage.getItem('emailOrgNotifiedIds') || '{}');
        const todayKey = getTodayFormatted();
        if (saved.date === todayKey) {
            notifiedAccountIds = new Set(saved.ids || []);
        } else {
            localStorage.removeItem('emailOrgNotifiedIds');
        }
    } catch (e) { localStorage.removeItem('emailOrgNotifiedIds'); }
})();
let selectedAccountIds = new Set();
let sortColumn = null;
let sortDirection = 'asc';
let loadRetryCount = 0;

// Error handling
window.addEventListener("error", (event) => {
    console.error("Runtime Error:", event.error);
});

function showError(message) {
    let errorBox = document.getElementById("errorBox");
    if (!errorBox) {
        errorBox = document.createElement("div");
        errorBox.id = "errorBox";
        document.body.appendChild(errorBox);
    }
    errorBox.innerText = message;
    setTimeout(() => { if (errorBox.parentNode) errorBox.remove(); }, 6000);
}

// Utilities
function extractDay(dateValue) {
    if (!dateValue) return null;
    const num = parseInt(dateValue, 10);
    if (!isNaN(num) && num >= 1 && num <= 31) return num;
    const d = new Date(dateValue);
    if (!isNaN(d.getDate())) return d.getDate();
    return null;
}

function formatDateForDisplay(dateValue) {
    const day = extractDay(dateValue);
    if (day === null) return '';
    return `Day ${day}`;
}

function formatDateForStorage(dayValue) {
    const day = extractDay(dayValue);
    if (day !== null) return String(day);
    return String(new Date().getDate());
}

function getTodayFormatted() {
    return String(new Date().getDate());
}

function getDaysUntilExpiry(dayValue) {
    const day = extractDay(dayValue);
    if (day === null) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const thisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const lastDayThisMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
    const clampedDay = Math.min(day, lastDayThisMonth);
    const thisMonthDate = new Date(today.getFullYear(), today.getMonth(), clampedDay);
    if (thisMonthDate >= today) {
        return Math.round((thisMonthDate - today) / (1000 * 60 * 60 * 24));
    }
    const lastDayNextMonth = new Date(today.getFullYear(), today.getMonth() + 2, 0).getDate();
    const clampedDayNext = Math.min(day, lastDayNextMonth);
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, clampedDayNext);
    return Math.round((nextMonth - today) / (1000 * 60 * 60 * 24));
}

function getDaysBadgeHTML(dateString) {
    const days = getDaysUntilExpiry(dateString);
    if (days === null) return '<span class="days-badge ok">-</span>';
    if (days < 0) return `<span class="days-badge expired">${days}d</span>`;
    if (days === 0) return '<span class="days-badge danger">Today</span>';
    if (days <= 7) return `<span class="days-badge warning">${days}d</span>`;
    return `<span class="days-badge ok">${days}d</span>`;
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
    setTimeout(() => messageEl.remove(), 3500);
}

function escapeHtml(text) {
    const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
    return String(text).replace(/[&<>"']/g, m => map[m]);
}

function escapeJS(text) {
    return String(text).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"').replace(/\n/g, '\\n').replace(/\r/g, '\\r');
}

function copyToClipboard(text) {
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text)
            .then(() => showMessage('Copied to clipboard'))
            .catch(() => fallbackCopy(text));
    } else {
        fallbackCopy(text);
    }
}

function fallbackCopy(text) {
    try {
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.left = '-999999px';
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showMessage('Copied to clipboard');
    } catch (err) {
        showMessage('Copy not supported', 'error');
    }
}

function validateEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validateDate(dateString) {
    if (!dateString) return true;
    const num = parseInt(dateString, 10);
    return !isNaN(num) && num >= 1 && num <= 31;
}

function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func(...args), wait);
    };
}

function toggleSection(id) {
    const el = document.getElementById(id);
    const arrow = document.getElementById(id.replace('Content', 'Arrow'));
    if (!el) return;
    const isHidden = el.style.display === 'none';
    el.style.display = isHidden ? '' : 'none';
    if (arrow) arrow.classList.toggle('collapsed', !isHidden);
}

// Dark Mode
function toggleDarkMode() {
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    document.body.setAttribute('data-theme', isDark ? '' : 'dark');
    localStorage.setItem('emailOrgDarkMode', isDark ? 'false' : 'true');
    document.getElementById('darkModeIcon').innerHTML = isDark ? '&#9789;' : '&#9788;';
}

function initDarkMode() {
    const isDark = localStorage.getItem('emailOrgDarkMode') === 'true';
    if (isDark) {
        document.body.setAttribute('data-theme', 'dark');
        document.getElementById('darkModeIcon').innerHTML = '&#9788;';
    }
}

// Notification Settings
function getNotificationSettings() {
    const defaults = { enabled: true, daysBefore: 30, notifyOnExpirationDay: true, notifyExpired: true, telegramEnabled: false, telegramBotToken: '', telegramChatId: '' };
    try {
        const saved = localStorage.getItem('emailOrgNotificationSettings');
        return saved ? { ...defaults, ...JSON.parse(saved) } : defaults;
    } catch { return defaults; }
}

function saveNotificationPrefs(settings) {
    localStorage.setItem('emailOrgNotificationSettings', JSON.stringify(settings));
}

function getExpiringAccounts() {
    const settings = getNotificationSettings();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const results = [];

    accounts.forEach(account => {
        if (!account.date) return;
        const days = getDaysUntilExpiry(account.date);

        if (settings.notifyExpired && days < 0) {
            results.push({ ...account, daysUntilExpiry: days, status: 'expired' });
        } else if (settings.notifyOnExpirationDay && days === 0) {
            results.push({ ...account, daysUntilExpiry: 0, status: 'today' });
        } else if (days !== null && days > 0 && days <= settings.daysBefore) {
            results.push({ ...account, daysUntilExpiry: days, status: 'expiring' });
        }
    });

    return results;
}

async function requestNotificationPermission() {
    if (!('Notification' in window)) return false;
    if (Notification.permission === 'granted') return true;
    if (Notification.permission === 'denied') return false;
    const result = await Notification.requestPermission();
    return result === 'granted';
}

function sendBrowserNotification(title, body) {
    if (Notification.permission !== 'granted') return;

    if ('serviceWorker' in navigator && navigator.serviceWorker.controller) {
        navigator.serviceWorker.ready.then(reg => {
            reg.showNotification(title, {
                body,
                tag: 'email-expiry-' + Date.now(),
                requireInteraction: true,
                icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">&#128231;</text></svg>'
            });
        }).catch(() => {
            try {
                const notif = new Notification(title, { body, tag: 'email-expiry-' + Date.now() });
                notif.onclick = () => { window.focus(); notif.close(); };
            } catch (e) { console.warn('Notification failed:', e); }
        });
    } else {
        try {
            const notif = new Notification(title, { body, tag: 'email-expiry-' + Date.now() });
            notif.onclick = () => { window.focus(); notif.close(); };
        } catch (e) {
            console.warn('Notification not available:', e);
            showMessage(`${title}: ${body}`, 'success');
        }
    }
}

async function sendTelegramMessage(text) {
    const settings = getNotificationSettings();
    if (!settings.telegramEnabled || !settings.telegramBotToken || !settings.telegramChatId) return;
    try {
        const url = `https://api.telegram.org/bot${settings.telegramBotToken}/sendMessage`;
        await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: settings.telegramChatId, text, parse_mode: 'HTML' })
        });
    } catch (e) {
        console.warn('Telegram notification failed:', e);
    }
}

async function testTelegramConnection() {
    const token = document.getElementById('telegramBotToken').value.trim();
    const chatId = document.getElementById('telegramChatId').value.trim();
    if (!token || !chatId) {
        showMessage('Please enter both Bot Token and Chat ID', 'error');
        return;
    }
    try {
        const url = `https://api.telegram.org/bot${token}/sendMessage`;
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: chatId, text: 'Email Account Organizer connected successfully!' })
        });
        if (res.ok) {
            showMessage('Telegram test message sent!');
        } else {
            const data = await res.json();
            showMessage('Telegram error: ' + (data.description || 'Unknown error'), 'error');
        }
    } catch (e) {
        showMessage('Failed to connect to Telegram', 'error');
    }
}

async function sendTestTelegram() {
    const settings = getNotificationSettings();
    if (!settings.telegramBotToken || !settings.telegramChatId) {
        showMessage('Set Bot Token and Chat ID first', 'error');
        return;
    }
    try {
        const url = `https://api.telegram.org/bot${settings.telegramBotToken}/sendMessage`;
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                chat_id: settings.telegramChatId,
                text: 'Test from Email Organizer!\nExpiry notifications are working.'
            })
        });
        if (res.ok) showMessage('Test Telegram sent!');
        else {
            const data = await res.json();
            showMessage('Telegram error: ' + (data.description || ''), 'error');
        }
    } catch (e) {
        showMessage('Failed to send test', 'error');
    }
}

function resetNotificationState() {
    notifiedAccountIds.clear();
    localStorage.setItem('emailOrgNotifiedIds', JSON.stringify({ date: getTodayFormatted(), ids: [] }));
    showMessage('Notification state reset. Reload or run checkAndNotify() to re-trigger.');
}

function checkAndNotify() {
    const settings = getNotificationSettings();
    const expiring = getExpiringAccounts();
    const newNotifs = expiring.filter(a => !notifiedAccountIds.has(a.id));
    if (newNotifs.length === 0) return;

    newNotifs.forEach(account => {
        notifiedAccountIds.add(account.id);
        localStorage.setItem('emailOrgNotifiedIds', JSON.stringify({ date: getTodayFormatted(), ids: [...notifiedAccountIds] }));
        let title, body;
        if (account.status === 'expired') {
            title = 'Email Account Past Due!';
            body = `${account.email} (${account.client}) - Day ${extractDay(account.date)} has passed.`;
        } else if (account.status === 'today') {
            title = 'Email Expires Today!';
            body = `${account.email} (${account.client}) - Day ${extractDay(account.date)} is today.`;
        } else {
            title = `Expires in ${account.daysUntilExpiry} day(s)`;
            body = `${account.email} (${account.client}) - Day ${extractDay(account.date)} in ${account.daysUntilExpiry} day(s).`;
        }
        if (settings.enabled) sendBrowserNotification(title, body);
        sendTelegramMessage(`<b>${title}</b>\n${body}`);
    });
    renderExpiringAccounts();
}

function openNotificationSettings() {
    const settings = getNotificationSettings();
    document.getElementById('notificationsEnabled').checked = settings.enabled;
    document.getElementById('notifyDaysBefore').value = settings.daysBefore;
    document.getElementById('notifyOnExpirationDay').checked = settings.notifyOnExpirationDay;
    document.getElementById('notifyExpired').checked = settings.notifyExpired;
    document.getElementById('telegramEnabled').checked = settings.telegramEnabled;
    document.getElementById('telegramBotToken').value = settings.telegramBotToken;
    document.getElementById('telegramChatId').value = settings.telegramChatId;
    document.getElementById('notificationSettingsModal').style.display = 'block';
}

function saveNotificationSettings() {
    const settings = {
        enabled: document.getElementById('notificationsEnabled').checked,
        daysBefore: Math.max(1, Math.min(365, parseInt(document.getElementById('notifyDaysBefore').value) || 1)),
        notifyOnExpirationDay: document.getElementById('notifyOnExpirationDay').checked,
        notifyExpired: document.getElementById('notifyExpired').checked,
        telegramEnabled: document.getElementById('telegramEnabled').checked,
        telegramBotToken: document.getElementById('telegramBotToken').value.trim(),
        telegramChatId: document.getElementById('telegramChatId').value.trim()
    };
    saveNotificationPrefs(settings);
    document.getElementById('notificationSettingsModal').style.display = 'none';

    if (settings.enabled) {
        requestNotificationPermission().then(granted => {
            if (granted) {
                showMessage('Notification settings saved');
                checkAndNotify();
            } else {
                settings.enabled = false;
                saveNotificationPrefs(settings);
                showMessage('Notifications blocked by browser', 'error');
            }
        });
    } else {
        showMessage('Notification settings saved');
    }
    renderExpiringAccounts();
    updateNotificationButton();
}

function updateNotificationButton() {
    const settings = getNotificationSettings();
    const btn = document.getElementById('notificationSettingsBtn');
    const text = document.getElementById('notifBtnText');
    if (settings.enabled) {
        btn.classList.add('active');
        text.textContent = 'Notifications ON';
    } else {
        btn.classList.remove('active');
        text.textContent = 'Notifications OFF';
    }
}

function startNotificationChecker() {
    if (notificationInterval) clearInterval(notificationInterval);
    checkAndNotify();
    notificationInterval = setInterval(checkAndNotify, 60 * 60 * 1000);
}

// Firebase CRUD
async function loadAccounts() {
    if (!db) {
        try {
            db = firebase.firestore();
            db.enablePersistence({ synchronizeTabs: true }).catch(() => {});
        } catch (e) {
            showError("Cannot connect to Firebase. Check internet.");
            return;
        }
    }
    try {
        showLoading();
        const snapshot = await db.collection('accounts').get();
        accounts = [];
        snapshot.forEach(doc => {
            const data = doc.data();
            accounts.push({
                id: doc.id,
                client: data.client || '',
                email: data.email || '',
                date: data.date || '',
                replacementEmail: data.replacementEmail || '',
                hasProblem: data.hasProblem || false,
                problemNote: data.problemNote || ''
            });
        });
        clients = [...new Set(accounts.map(a => a.client))].filter(Boolean);
        updateStats();
        renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
        updateNotificationButton();
        checkAndNotify();
        hideLoading();
    } catch (error) {
        console.error('Error loading accounts:', error);
        const msg = error.message || String(error);
        if (loadRetryCount < 3) {
            loadRetryCount++;
            console.log(`Retrying load (${loadRetryCount}/3)...`);
            hideLoading();
            setTimeout(() => loadAccounts(), 2000 * loadRetryCount);
            return;
        }
        loadRetryCount = 0;
        if (msg.includes('network') || msg.includes('fetch') || msg.includes('Failed')) {
            showError('Network error. Check your internet connection and try again.');
        } else {
            showError('Error loading accounts: ' + msg);
        }
        hideLoading();
    }
}

async function saveAccount(accountData) {
    const docRef = await db.collection('accounts').add({
        client: accountData.client,
        email: accountData.email,
        date: formatDateForStorage(accountData.date),
        replacementEmail: accountData.replacementEmail || '',
        hasProblem: accountData.hasProblem || false,
        problemNote: accountData.problemNote || ''
    });
    return docRef.id;
}

async function updateAccount(id, accountData) {
    await db.collection('accounts').doc(id).update({
        client: accountData.client,
        email: accountData.email,
        date: formatDateForStorage(accountData.date),
        replacementEmail: accountData.replacementEmail || '',
        hasProblem: accountData.hasProblem || false,
        problemNote: accountData.problemNote || ''
    });
}

async function deleteAccount(id) {
    await db.collection('accounts').doc(id).delete();
}

async function bulkSaveAccounts(accountsData) {
    const batch = db.batch();
    accountsData.forEach(d => {
        const ref = db.collection('accounts').doc();
        batch.set(ref, {
            client: d.client,
            email: d.email,
            date: formatDateForStorage(d.date),
            replacementEmail: '',
            hasProblem: false,
            problemNote: ''
        });
    });
    await batch.commit();
}

// Stats Dashboard
function updateStats() {
    let expired = 0, expiring = 0, problems = 0;
    accounts.forEach(a => {
        if (a.hasProblem) problems++;
        if (!a.date) return;
        const days = getDaysUntilExpiry(a.date);
        if (days !== null && days < 0) expired++;
        else if (days !== null && days <= 30) expiring++;
    });

    document.getElementById('statTotal').textContent = accounts.length;
    document.getElementById('statExpiring').textContent = expiring;
    document.getElementById('statExpired').textContent = expired;
    document.getElementById('statProblems').textContent = problems;
    document.getElementById('statClients').textContent = clients.length;
    document.getElementById('headerStats').textContent =
        `${accounts.length} accounts across ${clients.length} clients`;
}

// Selection
function toggleSelectAll(checkbox, context) {
    const checkboxes = document.querySelectorAll(`#${context === 'search' ? 'searchResultsBody' : context === 'expiring' ? 'expiringTableBody' : 'problemsTableBody'} input[type="checkbox"].row-select`);
    checkboxes.forEach(cb => {
        cb.checked = checkbox.checked;
        const id = cb.dataset.id;
        if (checkbox.checked) {
            selectedAccountIds.add(id);
            cb.closest('tr').classList.add('selected-row');
        } else {
            selectedAccountIds.delete(id);
            cb.closest('tr').classList.remove('selected-row');
        }
    });
    updateBulkBar();
}

function toggleRowSelect(checkbox, id) {
    if (checkbox.checked) {
        selectedAccountIds.add(id);
        checkbox.closest('tr').classList.add('selected-row');
    } else {
        selectedAccountIds.delete(id);
        checkbox.closest('tr').classList.remove('selected-row');
    }
    updateBulkBar();
}

function clearSelection() {
    selectedAccountIds.clear();
    document.querySelectorAll('.row-select').forEach(cb => {
        cb.checked = false;
        cb.closest('tr').classList.remove('selected-row');
    });
    document.querySelectorAll('#selectAllSearch, #selectAllExpiring, #selectAllProblems').forEach(cb => cb.checked = false);
    updateBulkBar();
}

function selectAllFiltered() {
    filteredAccounts.forEach(a => selectedAccountIds.add(a.id));
    renderSearchResults();
    updateBulkBar();
}

function updateBulkBar() {
    const bar = document.getElementById('bulkActionsBar');
    const count = selectedAccountIds.size;
    if (count > 0) {
        bar.style.display = 'flex';
        document.getElementById('bulkCount').textContent = count;
    } else {
        bar.style.display = 'none';
    }
}

// Bulk Actions
function getSelectedAccounts() {
    return accounts.filter(a => selectedAccountIds.has(a.id));
}

function bulkCopyEmails() {
    const selected = getSelectedAccounts();
    if (!selected.length) { showMessage('No accounts selected', 'error'); return; }
    copyToClipboard(selected.map(a => a.email).join('\n'));
}

function bulkExportSelected() {
    const selected = getSelectedAccounts();
    if (!selected.length) { showMessage('No accounts selected', 'error'); return; }
    const headers = ['Client', 'Email', 'Expiration Date', 'Replacement Email', 'Has Problem', 'Problem Note'];
    const csv = [
        headers.join(','),
        ...selected.map(a => [
            `"${a.client}"`, `"${a.email}"`, `"${formatDateForDisplay(a.date)}"`,
            `"${a.replacementEmail || ''}"`, `"${a.hasProblem ? 'Yes' : 'No'}"`, `"${a.problemNote || ''}"`
        ].join(','))
    ].join('\n');
    downloadFile(csv, `selected-accounts-${getTodayFormatted()}.csv`, 'text/csv');
    showMessage(`Exported ${selected.length} accounts`);
}

function bulkDeleteSelected() {
    const count = selectedAccountIds.size;
    if (!count) return;
    showConfirmModal(`Delete ${count} selected account(s)?`, async () => {
        try {
            showLoading();
            const batch = db.batch();
            selectedAccountIds.forEach(id => {
                batch.delete(db.collection('accounts').doc(id));
            });
            await batch.commit();
            accounts = accounts.filter(a => !selectedAccountIds.has(a.id));
            selectedAccountIds.clear();
            updateBulkBar();
            hideLoading();
            showMessage(`Deleted ${count} accounts`);
            clients = [...new Set(accounts.map(a => a.client))].filter(Boolean);
            loadRetryCount = 0;
            updateStats();
            renderAccounts();
            renderExpiringAccounts();
            renderProblemAccounts();
            if (isSearchActive) renderSearchResults();
        } catch (error) {
            hideLoading();
            showMessage('Error deleting accounts', 'error');
        }
    });
}

function openBulkEditModal() {
    if (!selectedAccountIds.size) return;
    document.getElementById('bulkEditDate').value = '';
    document.getElementById('bulkEditReplacement').value = '';
    document.getElementById('bulkEditClearProblems').checked = false;
    document.getElementById('bulkEditMarkProblems').checked = false;
    document.getElementById('bulkEditModal').style.display = 'block';
}

async function processBulkEdit() {
    const date = document.getElementById('bulkEditDate').value;
    const replacement = document.getElementById('bulkEditReplacement').value.trim();
    const clearProblems = document.getElementById('bulkEditClearProblems').checked;
    const markProblems = document.getElementById('bulkEditMarkProblems').checked;

    if (!date && !replacement && !clearProblems && !markProblems) {
        showMessage('No changes specified', 'error');
        return;
    }

    try {
        showLoading();
        const batch = db.batch();
        selectedAccountIds.forEach(id => {
            const account = accounts.find(a => a.id === id);
            if (!account) return;
            const updates = { ...account };
            if (date) updates.date = date;
            if (replacement) updates.replacementEmail = replacement;
            if (clearProblems) { updates.hasProblem = false; updates.problemNote = ''; }
            if (markProblems) { updates.hasProblem = true; }
            batch.update(db.collection('accounts').doc(id), {
                date: date ? updates.date : account.date,
                replacementEmail: replacement || account.replacementEmail,
                hasProblem: updates.hasProblem,
                problemNote: updates.problemNote || account.problemNote
            });
            Object.assign(account, updates);
        });
        await batch.commit();
        hideLoading();
        document.getElementById('bulkEditModal').style.display = 'none';
        showMessage(`Updated ${selectedAccountIds.size} accounts`);
        clearSelection();
        clients = [...new Set(accounts.map(a => a.client))].filter(Boolean);
        updateStats();
        renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
        if (isSearchActive) renderSearchResults();
    } catch (error) {
        hideLoading();
        showMessage('Error updating accounts', 'error');
    }
}

// Search
function performSearch() {
    const emailInput = document.getElementById('emailSearchInput').value.trim();
    const dateQuery = document.getElementById('dateSearchInput').value;
    const clientQuery = document.getElementById('clientSearchInput').value.trim().toLowerCase();
    const statusFilter = document.getElementById('statusFilter').value;

    if (!emailInput && !dateQuery && !clientQuery && !statusFilter) { clearSearch(); return; }

    const emailQueries = emailInput.split(/[\n,\s-]+/)
        .map(e => e.trim().toLowerCase())
        .filter(e => e && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));

    filteredAccounts = accounts.filter(account => {
        const emailMatch = !emailInput || emailQueries.some(q => account.email.toLowerCase().includes(q));
        const accountDay = extractDay(account.date);
        const searchDay = dateQuery ? parseInt(dateQuery, 10) : null;
        const dateMatch = !dateQuery || (accountDay !== null && accountDay === searchDay);
        const clientMatch = !clientQuery || account.client.toLowerCase().includes(clientQuery);

        let statusMatch = true;
        if (statusFilter === 'problem') statusMatch = account.hasProblem;
        else if (statusFilter === 'ok') statusMatch = !account.hasProblem;
        else if (statusFilter === 'expiring' || statusFilter === 'expired') {
            if (!account.date) statusMatch = false;
            else {
                const days = getDaysUntilExpiry(account.date);
                statusMatch = statusFilter === 'expired' ? days < 0 : (days !== null && days >= 0 && days <= 30);
            }
        }

        return emailMatch && dateMatch && clientMatch && statusMatch;
    });

    isSearchActive = true;
    renderSearchResults();
    if (filteredAccounts.length === 0) showMessage('No accounts found.', 'error');
}

function clearSearch() {
    document.getElementById('emailSearchInput').value = '';
    document.getElementById('dateSearchInput').value = '';
    document.getElementById('clientSearchInput').value = '';
    document.getElementById('statusFilter').value = '';
    document.getElementById('searchResults').style.display = 'none';
    isSearchActive = false;
    filteredAccounts = [];
}

function sortSearchResults(column) {
    if (sortColumn === column) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = column;
        sortDirection = 'asc';
    }

    filteredAccounts.sort((a, b) => {
        let valA = a[column] || '';
        let valB = b[column] || '';
        if (column === 'date') {
            valA = extractDay(a.date) || 9999;
            valB = extractDay(b.date) || 9999;
            return sortDirection === 'asc' ? valA - valB : valB - valA;
        }
        valA = valA.toLowerCase();
        valB = valB.toLowerCase();
        if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
        if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
        return 0;
    });

    renderSearchResults();
}

function renderSearchResults() {
    const container = document.getElementById('searchResults');
    const tbody = document.getElementById('searchResultsBody');
    const count = document.getElementById('resultCount');

    if (filteredAccounts.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="no-data">No accounts found.</td></tr>';
        count.textContent = '0 results';
        container.style.display = 'block';
        return;
    }

    const byClient = {};
    filteredAccounts.forEach(a => {
        if (!byClient[a.client]) byClient[a.client] = [];
        byClient[a.client].push(a);
    });

    tbody.innerHTML = Object.keys(byClient).sort().map(client => `
        <tr>
            <td colspan="8" class="client-header-row">
                <span class="client-name">${escapeHtml(client)} (${byClient[client].length})</span>
                <button class="btn btn-copy btn-small" onclick="copyToClipboard(decodeURIComponent(this.dataset.emails))" data-emails="${encodeURIComponent(byClient[client].map(a => a.email).join(', '))}">Copy All</button>
            </td>
        </tr>
        ${byClient[client].map(a => createAccountRowHTML(a, true)).join('')}
    `).join('');

    count.textContent = `${filteredAccounts.length} result${filteredAccounts.length !== 1 ? 's' : ''}`;
    container.style.display = 'block';
}

// Export
function exportToCSV() {
    const data = isSearchActive ? filteredAccounts : accounts;
    if (!data.length) { showMessage('No accounts to export', 'error'); return; }
    const headers = ['Client', 'Email', 'Expiration Date', 'Replacement Email', 'Has Problem', 'Problem Note'];
    const csv = [
        headers.join(','),
        ...data.map(a => [
            `"${a.client}"`, `"${a.email}"`, `"${formatDateForDisplay(a.date)}"`,
            `"${a.replacementEmail || ''}"`, `"${a.hasProblem ? 'Yes' : 'No'}"`, `"${a.problemNote || ''}"`
        ].join(','))
    ].join('\n');
    downloadFile(csv, `email-accounts-${getTodayFormatted()}.csv`, 'text/csv');
    showMessage(`CSV exported (${isSearchActive ? 'filtered' : 'all'})`);
}

function exportToExcel() {
    const data = isSearchActive ? filteredAccounts : accounts;
    if (!data.length) { showMessage('No accounts to export', 'error'); return; }
    const xml = `<?xml version="1.0"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
        <Worksheet ss:Name="Email Accounts"><Table>
            <Row>${['Client','Email','Expiration Date','Replacement Email','Has Problem','Problem Note'].map(h => `<Cell><Data ss:Type="String">${h}</Data></Cell>`).join('')}</Row>
            ${data.map(a => `<Row>
                <Cell><Data ss:Type="String">${a.client}</Data></Cell>
                <Cell><Data ss:Type="String">${a.email}</Data></Cell>
                <Cell><Data ss:Type="String">${formatDateForDisplay(a.date)}</Data></Cell>
                <Cell><Data ss:Type="String">${a.replacementEmail || ''}</Data></Cell>
                <Cell><Data ss:Type="String">${a.hasProblem ? 'Yes' : 'No'}</Data></Cell>
                <Cell><Data ss:Type="String">${a.problemNote || ''}</Data></Cell>
            </Row>`).join('')}
        </Table></Worksheet></Workbook>`;
    downloadFile(xml, `email-accounts-${getTodayFormatted()}.xls`, 'application/vnd.ms-excel');
    showMessage(`Excel exported (${isSearchActive ? 'filtered' : 'all'})`);
}

function exportToJSON() {
    const data = isSearchActive ? filteredAccounts : accounts;
    if (!data.length) { showMessage('No accounts to export', 'error'); return; }
    const json = JSON.stringify(data.map(a => ({
        client: a.client, email: a.email, expirationDate: formatDateForDisplay(a.date),
        rawDate: a.date, replacementEmail: a.replacementEmail || '',
        hasProblem: a.hasProblem, problemNote: a.problemNote || ''
    })), null, 2);
    downloadFile(json, `email-accounts-${getTodayFormatted()}.json`, 'application/json');
    showMessage(`JSON exported (${isSearchActive ? 'filtered' : 'all'})`);
}

function downloadFile(content, filename, type) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
}

// Rendering
function renderAccounts() {
    const container = document.getElementById('clientsContainer');
    container.innerHTML = '';
    const byClient = {};
    accounts.forEach(a => {
        if (!byClient[a.client]) byClient[a.client] = [];
        byClient[a.client].push(a);
    });
    Object.keys(byClient).sort().forEach(client => {
        container.appendChild(createClientSection(client, byClient[client]));
    });
    if (!Object.keys(byClient).length) {
        container.innerHTML = '<div class="no-data">No accounts found. Add your first client!</div>';
    }
}

function createClientSection(clientName, clientAccounts) {
    const isExpanded = expandedClients.has(clientName);
    const display = isExpanded ? clientAccounts : clientAccounts.slice(0, 5);
    const showToggle = clientAccounts.length > 5;

    const section = document.createElement('div');
    section.className = 'client-section';
    section.innerHTML = `
        <div class="client-header">
            <div class="client-name">${escapeHtml(clientName)} (${clientAccounts.length})</div>
            <div class="client-actions">
                <button class="btn btn-success btn-small" onclick="addNewAccount(decodeURIComponent(this.dataset.client))" data-client="${encodeURIComponent(clientName)}">+ Add</button>
                <button class="btn btn-primary btn-small" onclick="openBulkUpload(decodeURIComponent(this.dataset.client))" data-client="${encodeURIComponent(clientName)}">+ Bulk</button>
                <button class="btn btn-copy btn-small" onclick="copyToClipboard(decodeURIComponent(this.dataset.emails))" data-emails="${encodeURIComponent(clientAccounts.map(a => a.email).join(', '))}">Copy All</button>
            </div>
        </div>
        <div class="table-container" style="max-height:${showToggle && !isExpanded ? '300px' : 'none'};overflow-y:${showToggle && !isExpanded ? 'auto' : 'visible'};">
            <table>
                <thead>
                    <tr>
                        <th class="checkbox-col"><input type="checkbox" onchange="toggleClientSelect(this, '${escapeHtml(clientName)}')"></th>
                        <th>Email</th>
                        <th>Expiration</th>
                        <th>Replacement</th>
                        <th>Status</th>
                        <th>Days</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody id="client-${clientName.replace(/[^a-zA-Z0-9]/g, '-')}">
                    ${display.map(a => createAccountRowHTML(a)).join('')}
                </tbody>
            </table>
        </div>
        ${showToggle ? `<div class="toggle-container"><button class="btn btn-primary btn-small" onclick="toggleClientAccounts('${escapeHtml(clientName)}')">${isExpanded ? 'Show Less' : 'Show More'}</button></div>` : ''}
    `;
    return section;
}

function toggleClientSelect(checkbox, clientName) {
    const tbodyId = 'client-' + clientName.replace(/[^a-zA-Z0-9]/g, '-');
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.querySelectorAll('.row-select').forEach(cb => {
        cb.checked = checkbox.checked;
        const id = cb.dataset.id;
        if (checkbox.checked) {
            selectedAccountIds.add(id);
            cb.closest('tr').classList.add('selected-row');
        } else {
            selectedAccountIds.delete(id);
            cb.closest('tr').classList.remove('selected-row');
        }
    });
    updateBulkBar();
}

function toggleClientAccounts(clientName) {
    expandedClients.has(clientName) ? expandedClients.delete(clientName) : expandedClients.add(clientName);
    renderAccounts();
}

function createAccountRowHTML(account, isSearch) {
    const problemClass = account.hasProblem ? 'problem-row' : '';
    const isSelected = selectedAccountIds.has(account.id);
    const selectedClass = isSelected ? 'selected-row' : '';
    const statusHTML = account.hasProblem
        ? `<span class="status-badge problem" title="${escapeHtml(account.problemNote)}">Problem</span>`
        : `<span class="status-badge ok">OK</span>`;
    const replacementHTML = account.replacementEmail
        ? `<span class="replacement-email" title="Replacement set">${escapeHtml(account.replacementEmail)}</span>`
        : `<span class="no-replacement">-</span>`;
    const clientCell = isSearch ? `<td>${escapeHtml(account.client)}</td>` : '';

    return `
        <tr class="account-row ${problemClass} ${selectedClass}" data-id="${account.id}">
            <td class="checkbox-col"><input type="checkbox" class="row-select" data-id="${account.id}" ${isSelected ? 'checked' : ''} onchange="toggleRowSelect(this, '${account.id}')"></td>
            ${clientCell}
            <td class="email-cell">${escapeHtml(account.email)}</td>
            <td class="date-cell">${formatDateForDisplay(account.date)}</td>
            <td class="replacement-cell">${replacementHTML}</td>
            <td class="status-cell">${statusHTML}</td>
            <td>${getDaysBadgeHTML(account.date)}</td>
            <td class="actions-cell">
                <div class="action-buttons">
                    <button class="btn btn-primary btn-small" onclick="openReplacementModal('${account.id}')" title="Set replacement">&#9998;</button>
                    <button class="btn ${account.hasProblem ? 'btn-success' : 'btn-warning'} btn-small" onclick="openProblemModal('${account.id}')" title="${account.hasProblem ? 'Clear problem' : 'Mark problem'}">
                        ${account.hasProblem ? '&#10003;' : '&#9888;'}
                    </button>
                    <button class="btn btn-secondary btn-small" onclick="editAccount('${account.id}')" title="Edit">&#9998;</button>
                    <button class="btn btn-danger btn-small" onclick="confirmDeleteAccount('${account.id}')" title="Delete">&#10005;</button>
                    <button class="btn btn-copy btn-small" onclick="copyToClipboard(decodeURIComponent(this.dataset.email))" data-email="${encodeURIComponent(account.email)}" title="Copy email">&#128203;</button>
                </div>
            </td>
        </tr>`;
}

function renderExpiringAccounts() {
    const expiring = getExpiringAccounts();
    const tbody = document.getElementById('expiringTableBody');
    const noData = document.getElementById('noExpiringAccounts');
    const titleEl = document.getElementById('expiringSectionTitle');
    const settings = getNotificationSettings();

    let titleText = 'Accounts Expiring Soon';
    if (settings.daysBefore === 30 && settings.notifyOnExpirationDay) {
        titleText = 'Expiring Within 30 Days';
    } else if (settings.daysBefore === 1 && settings.notifyOnExpirationDay) {
        titleText = 'Expiring Tomorrow or Today';
    } else if (settings.daysBefore === 1) {
        titleText = 'Expiring Tomorrow';
    } else {
        titleText = `Expiring in ${settings.daysBefore} Days`;
    }
    if (settings.notifyExpired) titleText += ' (+ Expired)';
    titleEl.textContent = titleText;

    if (!expiring.length) {
        tbody.innerHTML = '';
        noData.style.display = 'block';
    } else {
        noData.style.display = 'none';
        tbody.innerHTML = expiring.map(a => {
            const row = createAccountRowHTML(a, true);
            if (a.status === 'expired') {
                return row.replace('class="account-row', 'class="account-row problem-row');
            }
            return row;
        }).join('');
    }
}

function renderProblemAccounts() {
    const problems = accounts.filter(a => a.hasProblem);
    const tbody = document.getElementById('problemsTableBody');
    const noData = document.getElementById('noProblemAccounts');

    if (!problems.length) {
        tbody.innerHTML = '';
        noData.style.display = 'block';
    } else {
        noData.style.display = 'none';
        tbody.innerHTML = problems.map(a => `
            <tr class="account-row problem-row" data-id="${a.id}">
                <td class="checkbox-col"><input type="checkbox" class="row-select" data-id="${a.id}" ${selectedAccountIds.has(a.id) ? 'checked' : ''} onchange="toggleRowSelect(this, '${a.id}')"></td>
                <td>${escapeHtml(a.client)}</td>
                <td>${escapeHtml(a.email)}</td>
                <td>${formatDateForDisplay(a.date)}</td>
                <td>${a.replacementEmail ? `<span class="replacement-email">${escapeHtml(a.replacementEmail)}</span>` : '-'}</td>
                <td><span class="problem-note">${escapeHtml(a.problemNote) || 'No details'}</span></td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn-success btn-small" onclick="openProblemModal('${a.id}')">Clear</button>
                        <button class="btn btn-copy btn-small" onclick="copyToClipboard(decodeURIComponent(this.dataset.email))" data-email="${encodeURIComponent(a.email)}">Copy</button>
                    </div>
                </td>
            </tr>
        `).join('');
    }
}

// Account CRUD UI
function addNewAccount(clientName) {
    const tbodyId = 'client-' + clientName.replace(/[^a-zA-Z0-9]/g, '-');
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    const row = document.createElement('tr');
    row.className = 'account-row editing-row';
    row.innerHTML = `
        <td class="checkbox-col"></td>
        <td><input type="email" class="new-email" placeholder="Enter email" required></td>
        <td><input type="number" class="new-date" min="1" max="31" value="${new Date().getDate()}" placeholder="Day"></td>
        <td>-</td>
        <td><span class="status-badge ok">OK</span></td>
        <td>-</td>
        <td>
            <div class="action-buttons">
                <button class="btn btn-success btn-small" onclick="saveNewAccount('${escapeHtml(clientName)}', this)">Save</button>
                <button class="btn btn-secondary btn-small" onclick="cancelNewAccount(this)">Cancel</button>
            </div>
        </td>`;
    tbody.insertBefore(row, tbody.firstChild);
    row.querySelector('.new-email').focus();
}

async function saveNewAccount(clientName, button) {
    const row = button.closest('tr');
    const email = row.querySelector('.new-email').value.trim();
    const date = row.querySelector('.new-date').value;

    if (!email || !validateEmail(email)) { showMessage('Please enter a valid email', 'error'); return; }
    if (!validateDate(date)) { showMessage('Please enter a valid date', 'error'); return; }

    try {
        showLoading();
        const data = { client: clientName, email, date: date || String(new Date().getDate()), replacementEmail: '', hasProblem: false, problemNote: '' };
        const id = await saveAccount(data);
        accounts.push({ id, ...data });
        hideLoading();
        showMessage('Account added');
        clients = [...new Set(accounts.map(a => a.client))].filter(Boolean);
        updateStats();
        renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
        if (isSearchActive) performSearch();
    } catch (error) {
        hideLoading();
        showMessage('Error adding account', 'error');
    }
}

function cancelNewAccount(button) { button.closest('tr').remove(); }

function editAccount(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) return;

    let row = null;
    for (const r of document.querySelectorAll('tr.account-row')) {
        if (r.dataset.id === id) { row = r; break; }
    }
    if (!row) return;

    const cells = Array.from(row.querySelectorAll('td'));
    let emailCell, dateCell, actionsCell;

    const clientOffset = isSearchActive ? 1 : 0;
    if (cells.length >= 5 + clientOffset) {
        emailCell = cells[1 + clientOffset];
        dateCell = cells[2 + clientOffset];
        actionsCell = cells[cells.length - 1];
    }

    if (!emailCell || !dateCell || !actionsCell) return;

    emailCell.innerHTML = `<input type="email" class="edit-email" value="${escapeHtml(account.email)}">`;
    dateCell.innerHTML = `<input type="number" class="edit-date" min="1" max="31" value="${extractDay(account.date) || ''}" placeholder="Day">`;
    actionsCell.innerHTML = `
        <div class="action-buttons">
            <button class="btn btn-success btn-small" onclick="saveAccountEdit('${id}')">Save</button>
            <button class="btn btn-secondary btn-small" onclick="cancelAccountEdit('${id}')">Cancel</button>
        </div>`;
    row.classList.add('editing-row');
    row.querySelector('.edit-email').focus();
}

async function saveAccountEdit(id) {
    let row = document.querySelector(`tr.account-row[data-id="${id}"]`);
    if (!row) return;

    const email = row.querySelector('.edit-email').value.trim();
    const date = row.querySelector('.edit-date').value;

    if (!email || !validateEmail(email)) { showMessage('Valid email required', 'error'); return; }
    if (!validateDate(date)) { showMessage('Valid date required', 'error'); return; }

    try {
        showLoading();
        const account = accounts.find(a => a.id === id);
        const updated = {
            client: account.client, email, date: date || String(new Date().getDate()),
            replacementEmail: account.replacementEmail, hasProblem: account.hasProblem, problemNote: account.problemNote
        };
        await updateAccount(id, updated);
        Object.assign(account, updated);
        hideLoading();
        showMessage('Account updated');
        updateStats();
        if (isSearchActive) renderSearchResults(); else renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
    } catch (error) {
        hideLoading();
        showMessage('Error updating', 'error');
    }
}

function cancelAccountEdit(id) {
    if (isSearchActive) renderSearchResults(); else renderAccounts();
}

// Delete
function confirmDeleteAccount(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) return;
    showConfirmModal(`Delete ${account.email}?`, () => deleteAccountConfirmed(id));
}

async function deleteAccountConfirmed(id) {
    try {
        showLoading();
        await deleteAccount(id);
        accounts = accounts.filter(a => a.id !== id);
        selectedAccountIds.delete(id);
        clients = [...new Set(accounts.map(a => a.client))].filter(Boolean);
        hideLoading();
        showMessage('Account deleted');
        updateStats();
        updateBulkBar();
        renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
        if (isSearchActive) performSearch();
    } catch (error) {
        hideLoading();
        showMessage('Error deleting', 'error');
    }
}

// Replacement Modal
function openReplacementModal(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) return;
    currentReplacementAccountId = id;
    document.getElementById('replacementCurrentEmail').textContent = `Current: ${account.email}`;
    document.getElementById('replacementEmailInput').value = account.replacementEmail || '';
    document.getElementById('replacementModal').style.display = 'block';
    document.getElementById('replacementEmailInput').focus();
}

async function saveReplacement() {
    const id = currentReplacementAccountId;
    const account = accounts.find(a => a.id === id);
    if (!account) return;
    const newEmail = document.getElementById('replacementEmailInput').value.trim();
    if (newEmail && !validateEmail(newEmail)) { showMessage('Invalid email format', 'error'); return; }

    try {
        showLoading();
        account.replacementEmail = newEmail;
        await updateAccount(id, account);
        hideLoading();
        document.getElementById('replacementModal').style.display = 'none';
        showMessage(newEmail ? 'Replacement email saved' : 'Replacement cleared');
        if (isSearchActive) renderSearchResults(); else renderAccounts();
        renderExpiringAccounts();
    } catch (error) {
        hideLoading();
        showMessage('Error saving replacement', 'error');
    }
}

function clearReplacement() {
    document.getElementById('replacementEmailInput').value = '';
    saveReplacement();
}

// Problem Modal
function openProblemModal(id) {
    const account = accounts.find(a => a.id === id);
    if (!account) return;
    currentProblemAccountId = id;
    document.getElementById('problemCurrentEmail').textContent = `Account: ${account.email}`;
    document.getElementById('problemNoteInput').value = account.problemNote || '';
    document.getElementById('problemModal').style.display = 'block';
    document.getElementById('problemNoteInput').focus();
}

async function saveProblem() {
    const id = currentProblemAccountId;
    const account = accounts.find(a => a.id === id);
    if (!account) return;
    const note = document.getElementById('problemNoteInput').value.trim();

    try {
        showLoading();
        account.hasProblem = true;
        account.problemNote = note;
        await updateAccount(id, account);
        hideLoading();
        document.getElementById('problemModal').style.display = 'none';
        showMessage('Account marked as problem');
        updateStats();
        if (isSearchActive) renderSearchResults(); else renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
    } catch (error) {
        hideLoading();
        showMessage('Error marking problem', 'error');
    }
}

async function clearProblem() {
    const id = currentProblemAccountId;
    const account = accounts.find(a => a.id === id);
    if (!account) return;

    try {
        showLoading();
        account.hasProblem = false;
        account.problemNote = '';
        await updateAccount(id, account);
        hideLoading();
        document.getElementById('problemModal').style.display = 'none';
        showMessage('Problem cleared');
        updateStats();
        if (isSearchActive) renderSearchResults(); else renderAccounts();
        renderExpiringAccounts();
        renderProblemAccounts();
    } catch (error) {
        hideLoading();
        showMessage('Error clearing problem', 'error');
    }
}

// Bulk Upload
function openBulkUpload(clientName) {
    currentBulkClient = clientName;
    document.getElementById('bulkEmailInput').value = '';
    document.getElementById('bulkUploadModal').style.display = 'block';
    document.getElementById('bulkEmailInput').focus();
}

async function processBulkUpload() {
    const input = document.getElementById('bulkEmailInput').value.trim();
    if (!input) { showMessage('Enter some emails', 'error'); return; }

    const lines = input.split('\n').filter(l => l.trim());
    const toAdd = [];
    const errors = [];

    lines.forEach((line, i) => {
        const parts = line.trim().split(',').map(p => p.trim());
        const email = parts[0];
        let day = parts[1] || String(new Date().getDate());

        if (!email || !email.includes('@')) { errors.push(`Line ${i + 1}: Invalid email`); return; }
        const dayNum = parseInt(day, 10);
        if (isNaN(dayNum) || dayNum < 1 || dayNum > 31) {
            errors.push(`Line ${i + 1}: Invalid day (use 1-31)`); return;
        }
        toAdd.push({ client: currentBulkClient, email, date: String(dayNum) });
    });

    if (errors.length) { showMessage(errors.join('\n'), 'error'); return; }
    if (!toAdd.length) { showMessage('No valid emails', 'error'); return; }

    try {
        showLoading();
        await bulkSaveAccounts(toAdd);
        await loadAccounts();
        hideLoading();
        document.getElementById('bulkUploadModal').style.display = 'none';
        showMessage(`Added ${toAdd.length} accounts`);
    } catch (error) {
        hideLoading();
        showMessage('Error uploading', 'error');
    }
}

// Client Management
function showAddClientForm() {
    document.getElementById('addClientBtn').style.display = 'none';
    document.getElementById('newClientForm').style.display = 'flex';
    document.getElementById('newClientName').focus();
}

function hideAddClientForm() {
    document.getElementById('addClientBtn').style.display = 'inline-flex';
    document.getElementById('newClientForm').style.display = 'none';
    document.getElementById('newClientName').value = '';
}

async function saveNewClient() {
    const name = document.getElementById('newClientName').value.trim();
    if (!name) { showMessage('Enter a client name', 'error'); return; }
    if (clients.includes(name)) { showMessage('Client already exists', 'error'); return; }

    try {
        showLoading();
        const placeholder = { client: name, email: 'example@email.com', date: String(new Date().getDate()), replacementEmail: '', hasProblem: false, problemNote: '' };
        const id = await saveAccount(placeholder);
        accounts.push({ id, ...placeholder });
        clients.push(name);
        hideLoading();
        hideAddClientForm();
        showMessage(`Client "${name}" added`);
        updateStats();
        renderAccounts();
        setTimeout(() => editAccount(id), 100);
    } catch (error) {
        hideLoading();
        showMessage('Error adding client', 'error');
    }
}

// Modal Helpers
let _confirmCallback = null;

function showConfirmModal(message, onConfirm) {
    const modal = document.getElementById('confirmModal');
    document.getElementById('confirmMessage').textContent = message;
    _confirmCallback = onConfirm;
    modal.style.display = 'block';
}

// Event Listeners
document.addEventListener('DOMContentLoaded', function() {
    try {
        if (typeof firebase !== 'undefined' && !firebase.apps.length) {
            firebase.initializeApp(firebaseConfig);
        }
        db = firebase.firestore();
        db.enablePersistence({ synchronizeTabs: true }).catch(() => {});
    } catch (error) {
        console.error("Firebase init failed:", error);
        showError("Failed to connect to Firebase.");
    }

    initDarkMode();
    updateNotificationButton();
    loadAccounts();
    if (getNotificationSettings().enabled) requestNotificationPermission();
    startNotificationChecker();

    document.getElementById('darkModeBtn').addEventListener('click', toggleDarkMode);
    document.getElementById('addClientBtn').addEventListener('click', showAddClientForm);
    document.getElementById('saveClientBtn').addEventListener('click', saveNewClient);
    document.getElementById('cancelClientBtn').addEventListener('click', hideAddClientForm);
    document.getElementById('newClientName').addEventListener('keypress', e => { if (e.key === 'Enter') saveNewClient(); });

    document.getElementById('searchBtn').addEventListener('click', performSearch);
    document.getElementById('clearSearchBtn').addEventListener('click', clearSearch);
    ['emailSearchInput', 'dateSearchInput', 'clientSearchInput', 'statusFilter'].forEach(id => {
        const el = document.getElementById(id);
        el.addEventListener('input', debounce(performSearch, 300));
        el.addEventListener('change', debounce(performSearch, 300));
        el.addEventListener('keypress', e => { if (e.key === 'Enter') performSearch(); });
    });

    document.getElementById('exportCSVBtn').addEventListener('click', exportToCSV);
    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
    document.getElementById('exportJSONBtn').addEventListener('click', exportToJSON);

    document.getElementById('processBulkUpload').addEventListener('click', processBulkUpload);
    document.getElementById('cancelBulkUpload').addEventListener('click', () => { document.getElementById('bulkUploadModal').style.display = 'none'; });

    document.getElementById('saveReplacementBtn').addEventListener('click', saveReplacement);
    document.getElementById('clearReplacementBtn').addEventListener('click', clearReplacement);
    document.getElementById('cancelReplacementBtn').addEventListener('click', () => { document.getElementById('replacementModal').style.display = 'none'; });
    document.getElementById('replacementEmailInput').addEventListener('keypress', e => { if (e.key === 'Enter') saveReplacement(); });

    document.getElementById('saveProblemBtn').addEventListener('click', saveProblem);
    document.getElementById('clearProblemBtn').addEventListener('click', clearProblem);
    document.getElementById('cancelProblemBtn').addEventListener('click', () => { document.getElementById('problemModal').style.display = 'none'; });

    document.querySelectorAll('.close').forEach(btn => {
        btn.addEventListener('click', function() { this.closest('.modal').style.display = 'none'; });
    });

    document.getElementById('confirmYes').addEventListener('click', () => {
        document.getElementById('confirmModal').style.display = 'none';
        if (_confirmCallback) { _confirmCallback(); _confirmCallback = null; }
    });
    document.getElementById('confirmNo').addEventListener('click', () => {
        document.getElementById('confirmModal').style.display = 'none';
        _confirmCallback = null;
    });

    window.addEventListener('click', e => {
        if (e.target.classList.contains('modal')) e.target.style.display = 'none';
    });

    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') document.querySelectorAll('.modal').forEach(m => m.style.display = 'none');
        if (e.ctrlKey && e.key === 'n') { e.preventDefault(); showAddClientForm(); }
        if (e.ctrlKey && e.key === 'f') { e.preventDefault(); document.getElementById('emailSearchInput').focus(); }
        if (e.ctrlKey && e.key === 'e') { e.preventDefault(); exportToCSV(); }
        if (e.ctrlKey && e.key === 'a' && !e.target.matches('input, textarea, select')) {
            e.preventDefault();
            accounts.forEach(a => selectedAccountIds.add(a.id));
            renderAccounts();
            if (isSearchActive) renderSearchResults();
            updateBulkBar();
        }
    });
});
