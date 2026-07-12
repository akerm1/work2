const admin = require('firebase-admin');

// Initialize Firebase Admin
const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
});
const db = admin.firestore();

// Telegram config from env
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;

// Notification settings (defaults matching your web app)
const DAYS_BEFORE = parseInt(process.env.NOTIFY_DAYS_BEFORE || '30', 10);
const NOTIFY_EXPIRED = (process.env.NOTIFY_EXPIRED || 'true') === 'true';
const NOTIFY_EXPIRATION_DAY = (process.env.NOTIFY_EXPIRATION_DAY || 'true') === 'true';

function extractDay(dateValue) {
    if (!dateValue) return null;
    const num = parseInt(dateValue, 10);
    if (!isNaN(num) && num >= 1 && num <= 31) return num;
    return null;
}

function getDaysUntilExpiry(dayValue) {
    const day = extractDay(dayValue);
    if (day === null) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const thisMonth = new Date(today.getFullYear(), today.getMonth(), day);
    if (thisMonth >= today) {
        return Math.round((thisMonth - today) / (1000 * 60 * 60 * 24));
    }
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, day);
    return Math.round((nextMonth - today) / (1000 * 60 * 60 * 24));
}

function getTodayKey() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

async function sendTelegramMessage(text) {
    if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) {
        console.log('Telegram not configured, skipping message');
        return false;
    }
    try {
        const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text, parse_mode: 'HTML' }),
        });
        if (!res.ok) {
            const data = await res.json();
            console.error('Telegram API error:', data.description || 'Unknown');
            return false;
        }
        return true;
    } catch (e) {
        console.error('Telegram send failed:', e.message);
        return false;
    }
}

async function wasNotifiedToday(accountId, todayKey) {
    const doc = await db.collection('notificationLog').doc(accountId).get();
    if (!doc.exists) return false;
    return doc.data().lastNotified === todayKey;
}

async function markNotified(accountId, todayKey) {
    await db.collection('notificationLog').doc(accountId).set({
        lastNotified: todayKey,
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
}

async function main() {
    console.log(`[${new Date().toISOString()}] Starting expiry check...`);
    console.log(`Settings: DAYS_BEFORE=${DAYS_BEFORE}, NOTIFY_EXPIRED=${NOTIFY_EXPIRED}, NOTIFY_EXPIRATION_DAY=${NOTIFY_EXPIRATION_DAY}`);

    const todayKey = getTodayKey();
    const snapshot = await db.collection('accounts').get();
    const accounts = [];
    snapshot.forEach(doc => {
        const data = doc.data();
        accounts.push({
            id: doc.id,
            client: data.client || '',
            email: data.email || '',
            date: data.date || '',
        });
    });

    console.log(`Found ${accounts.length} accounts in Firestore`);

    let sentCount = 0;
    let skippedCount = 0;

    for (const account of accounts) {
        if (!account.date) continue;

        const days = getDaysUntilExpiry(account.date);
        if (days === null) continue;

        let shouldNotify = false;
        let title = '';
        let body = '';

        if (NOTIFY_EXPIRED && days < 0) {
            shouldNotify = true;
            title = 'Email Account Past Due!';
            body = `${account.email} (${account.client}) - Day ${extractDay(account.date)} has passed.`;
        } else if (NOTIFY_EXPIRATION_DAY && days === 0) {
            shouldNotify = true;
            title = 'Email Expires Today!';
            body = `${account.email} (${account.client}) - Day ${extractDay(account.date)} is today.`;
        } else if (days > 0 && days <= DAYS_BEFORE) {
            shouldNotify = true;
            title = `Expires in ${days} day(s)`;
            body = `${account.email} (${account.client}) - Day ${extractDay(account.date)} in ${days} day(s).`;
        }

        if (!shouldNotify) continue;

        const alreadyNotified = await wasNotifiedToday(account.id, todayKey);
        if (alreadyNotified) {
            skippedCount++;
            continue;
        }

        const message = `<b>${title}</b>\n${body}`;
        const sent = await sendTelegramMessage(message);
        if (sent) {
            await markNotified(account.id, todayKey);
            sentCount++;
            console.log(`Sent: ${account.email} - ${title}`);
        }
    }

    console.log(`Done. Sent: ${sentCount}, Skipped (already notified): ${skippedCount}`);
    await admin.app().delete();
}

main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
