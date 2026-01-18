// index.js - Enhanced with Google Sheets Integration

const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const express = require('express');
const fs = require('fs');
const googleSheetsService = require('./services/googleSheets.service');
require('dotenv').config({ path: './key/.env' });


const app = express();
app.use(express.json());

app.get('/', (req, res) => {
    res.json({
        message: 'Glam&Glow WhatsApp Bot',
        status: 'Running',
        timestamp: new Date().toISOString()
    });
});

app.get('/test', (req, res) => {
    res.json({
        message: 'Test endpoint working! 🎉',
        serverTime: new Date().toISOString()
    });
});

app.get('/health', (req, res) => {
    res.json({
        status: 'OK',
        uptime: process.uptime(),
        timestamp: new Date().toISOString()
    });
});

// Log all requests
app.use((req, res, next) => {
    console.log('📨 Request:', req.method, req.url);
    next();
});

// 📁 Storage files
const STORAGE_FILE = './customer_phones.json';
const WEBHOOK_LOG_FILE = './webhook_logs.json';

// 👑 MULTIPLE STORE OWNERS
const STORE_OWNERS = [
    '201001933044', // 👈 Owner 1 - Replace with actual number

];

// Initialize storage
let customerPhones = {};
let webhookLogs = {};

// Load existing data
function loadStorageData() {
    try {
        if (fs.existsSync(STORAGE_FILE)) {
            const data = fs.readFileSync(STORAGE_FILE, 'utf8');
            customerPhones = JSON.parse(data);
            console.log(`📋 Loaded ${Object.keys(customerPhones).length} customer phones`);
        }

        if (fs.existsSync(WEBHOOK_LOG_FILE)) {
            const data = fs.readFileSync(WEBHOOK_LOG_FILE, 'utf8');
            webhookLogs = JSON.parse(data);
            console.log(`📊 Loaded ${Object.keys(webhookLogs).length} webhook logs`);
        }
    } catch (error) {
        console.log('📋 Starting with fresh storage');
        customerPhones = {};
        webhookLogs = {};
    }
}

function saveStorageData() {
    try {
        fs.writeFileSync(STORAGE_FILE, JSON.stringify(customerPhones, null, 2));
        fs.writeFileSync(WEBHOOK_LOG_FILE, JSON.stringify(webhookLogs, null, 2));
    } catch (error) {
        console.error('❌ Error saving storage data:', error);
    }
}

loadStorageData();

// WhatsApp Client
const client = new Client({
    authStrategy: new LocalAuth({
        clientId: "glamglow-bot",
        dataPath: "./sessions"
    }),
    puppeteer: {
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

client.on('qr', (qr) => {
    console.log('📱 Scan this QR code:');
    qrcode.generate(qr, { small: true });
});

client.on('ready', async () => {
    console.log('✅ WhatsApp client is ready!');

    // Initialize Google Sheets
    await googleSheetsService.initialize();
});

client.initialize();
// 🆕 WHATSAPP MESSAGE LISTENER - Handle Customer Confirmations
// 🆕 WHATSAPP MESSAGE LISTENER - Handle Customer Confirmations
client.on('message', async (message) => {
    try {
        // Ignore group messages
        const chat = await message.getChat();
        if (chat.isGroup) {
            return;
        }

        // Get sender phone
        const senderPhone = message.from.split('@')[0];

        // Ignore store owners
        if (STORE_OWNERS.includes(senderPhone)) {
            console.log(`👑 Ignoring message from store owner: ${senderPhone}`);
            return;
        }

        // 🔒 HARD STOP — already confirmed (from customer_phones.json)
        const alreadyConfirmed = Object.values(customerPhones)
            .some(c => c.phone === senderPhone && c.confirmed === true);

        if (alreadyConfirmed) {
            console.log(`⛔ Message ignored — order already confirmed for ${senderPhone}`);
            return;
        }

        const messageText = message.body.trim().toLowerCase();
        console.log(`💬 Customer message from ${senderPhone}: "${messageText}"`);

        // Find pending order
        const pendingOrder = await googleSheetsService.getPendingOrderByPhone(senderPhone);

        if (!pendingOrder) {
            console.log(`ℹ️ No pending order found for ${senderPhone}`);
            return;
        }

        // Store last customer message (ONLY BEFORE confirmation)
        await googleSheetsService.updateLastMessage(
            pendingOrder.orderId,
            message.body
        );
        console.log(`💾 Stored last message for order ${pendingOrder.orderNumber}`);

        // Confirmation keywords
        const confirmationKeywords = [
            'yes', 'confirm', 'confirmed',
            'ok', 'okay',
            'تم', 'موافق', 'تأكيد', 'نعم'
        ];

        const isConfirmation = confirmationKeywords.some(keyword =>
            messageText.includes(keyword)
        );

        if (!isConfirmation) {
            console.log(`ℹ️ Message is not a confirmation keyword`);
            return;
        }

        console.log(`✅ Confirmation detected for order ${pendingOrder.orderNumber}`);

        // 🔒 Mark order as confirmed in customer_phones.json
        if (customerPhones[pendingOrder.orderId]) {
            customerPhones[pendingOrder.orderId].confirmed = true;
            customerPhones[pendingOrder.orderId].confirmedAt = new Date().toISOString();
            saveStorageData();
        }

        // Update order status in Google Sheets
        await googleSheetsService.updateOrderStatus(
            pendingOrder.orderId,
            'CONFIRMED',
            { confirmedAt: new Date().toISOString() }
        );

        // Thank customer
        const thankYouMessage = `Thank you ${pendingOrder.customerName}! 💛

Your order #${pendingOrder.orderNumber} has been confirmed and will be processed shortly! 🎉

We'll keep you updated on your order status.

Thanks for choosing Glam&Glow ✨`;

        await sendWhatsApp(
            senderPhone,
            thankYouMessage,
            pendingOrder.customerName
        );

        // Notify store owners
        const ownerNotification = `✅ Order Confirmed by Customer!

Order #${pendingOrder.orderNumber}
Customer: ${pendingOrder.customerName}
Phone: ${senderPhone}

Customer's message: "${message.body}"

The order is now confirmed and ready for processing.`;

        for (const ownerPhone of STORE_OWNERS) {
            await sendWhatsApp(
                ownerPhone,
                ownerNotification,
                `Owner ${STORE_OWNERS.indexOf(ownerPhone) + 1}`
            );
        }

        console.log(`🎉 Order ${pendingOrder.orderNumber} confirmed successfully`);

    } catch (error) {
        console.error('❌ Error handling customer message:', error);
    }
});



// 🧾 Debug Middleware - log every incoming request in detail
app.use((req, res, next) => {
    console.log('📨 New Request Received:');
    console.log('➡️ Method:', req.method);
    console.log('➡️ URL:', req.originalUrl);
    console.log('➡️ Headers:', req.headers);

    // Try to print parsed body if available
    if (req.body && Object.keys(req.body).length > 0) {
        console.log('📦 Parsed Body:', JSON.stringify(req.body, null, 2));
        return next();
    }

    // If body is empty, print raw data to debug webhook format
    let rawData = '';
    req.on('data', chunk => rawData += chunk);
    req.on('end', () => {
        if (rawData.length > 0) {
            console.log('⚠️ Raw Body (unparsed):', rawData);
            try {
                const parsed = JSON.parse(rawData);
                console.log('✅ Parsed manually:', JSON.stringify(parsed, null, 2));
            } catch {
                console.log('❌ Failed to parse JSON manually');
            }
        } else {
            console.log('⚠️ No body received in request.');
        }
    });

    next();
});

// 🎯 MAIN WEBHOOK ENDPOINT
app.post('/webhooks/wuilt', async (req, res) => {
    const startTime = Date.now();

    try {
        console.log('📦 Received Webhook');
        console.log('📨 Full Body:', JSON.stringify(req.body, null, 2));

        // ✅ Adjusted for new Wuilt format
        const data = req.body.data || {};
        const event = data.event || req.body.event;
        const payload = data.payload || req.body.payload;
        const metadata = data.metadata || req.body.metadata;

        // ⬇️ دائماً نرد بـ 200 أولاً
        res.status(200).json({
            status: 'OK',
            message: 'Webhook received successfully',
            timestamp: new Date().toISOString()
        });

        // ⬇️ ثم نتحقق إذا كان فيه بيانات صحيحة عشان نكمل
        if (!event || !payload) {
            console.log('⚠️ Invalid webhook format - skipping processing');
            return;
        }

        // Duplicate detection
        const webhookSignature = `${event}-${metadata?.timestamp}-${payload.order?.orderId || payload.order?._id}`;

        //if (webhookLogs[webhookSignature]) {
        //   console.log('🔄 Duplicate webhook skipped:', webhookSignature);
        //   return;
        //}

        webhookLogs[webhookSignature] = {
            processedAt: new Date().toISOString(),
            event: event
        };
        saveStorageData();

        if (!client.info) {
            console.log('⚠️ WhatsApp client not ready');
            return;
        }

        console.log(`🔔 Processing ${event}`);

        // ✅ Print basic order info for debugging
        if (payload?.order) {
            const order = payload.order;
            console.log('🆔 Order ID:', order._id || order.orderId);
            console.log('👤 Customer:', order.customer?.name, order.customer?.phone || order.shippingAddress?.phone);
            console.log('💰 Total:', order.totalPrice?.amount, order.totalPrice?.currencyCode);
            console.log('📦 Items Count:', order.items?.length);
            if (order.items?.length) {
                console.log('🛍️ Items List:');
                order.items.forEach(item => {
                    console.log(`   - ${item.title} × ${item.quantity}`);
                });
            }
        }

        // Process webhook
        setImmediate(async () => {
            try {
                let result;
                switch (event) {
                    case 'ORDER_PLACED':
                        result = await handleOrderPlaced(payload.order);
                        break;
                    case 'SHIPMENT_UPDATED':
                        result = await handleShipmentUpdate(payload);
                        break;
                    case 'ORDER_CANCELED':
                        result = await handleOrderCancel(payload.order);
                        break;
                    default:
                        console.log(`⚡ Unhandled event: ${event}`);
                        result = 'unhandled_event';
                }

                console.log(`✅ ${event} processed - Result: ${result}`);

            } catch (error) {
                console.error(`❌ Error processing ${event}:`, error);
            }
        });

    } catch (error) {
        console.error('💥 Webhook error:', error);
        // حتى في حالة الخطأ نكون قد أرسلنا 200 بالفعل
    }
});

// 🛍️ HANDLE NEW ORDER - WITH MULTIPLE OWNERS + GOOGLE SHEETS
async function handleOrderPlaced(order) {
    try {
        if (!order?.customer || !order?.shippingAddress) {
            return 'invalid_data';
        }

        const orderId = order._id;
        const customerName = order.customer.name;
        const rawPhone = order.shippingAddress.phone;
        const customerPhone = formatPhone(rawPhone);

        if (!customerPhone) {
            return 'invalid_phone';
        }

        // Store customer phone (EXISTING BEHAVIOR - DO NOT CHANGE)
        storeCustomerPhone(orderId, customerPhone, customerName);

        const orderNumber = order.orderSerial;
        const totalAmount = order.totalPrice.amount;

        // 🆕 ADD ORDER TO GOOGLE SHEETS
        await googleSheetsService.addOrder(order);

        // 🕯️ CUSTOMER MESSAGE - Order Confirmation (EXISTING - DO NOT CHANGE)
        const customerMessage = `Hey ${customerName} 💛

Your Glam&Glow order (#${orderNumber}) has been successfully placed and it's currently under processing! 🕯

🧾 Total: ${totalAmount} EGP

You'll get another message once the shipping company picks it up for delivery 🚚

✨ *Important*:
• To confirm and proceed with your order, please reply to this message with "*Confirmed*".
• If no confirmation message is received within 24 hours, the order will be automatically cancelled.

Thanks for choosing Glam&Glow — we can't wait for you to enjoy your order! ✨`;

        // 🛍️ STORE OWNERS MESSAGE - New Order Notification (EXISTING - DO NOT CHANGE)
        const itemsList = order.items.map(item => {
            const sizeOption = item.options.find(opt => opt.name === 'Size');
            const size = sizeOption ? sizeOption.value : '';
            return `- ${item.title}${size ? ` (${size})` : ''} × ${item.quantity}`;
        }).join('\n');

        const ownerMessage = `🛍 New Order Received!

👤 Customer: ${order.customer.name}  
📞 Phone: ${order.customer.phone}  

🧾 Order Number: #${orderNumber}  
💰 Total: ${totalAmount} ${order.totalPrice.currencyCode}  
🚚 Shipping: ${order.shippingRateName}  

📦 Items (${order.itemsCount || order.items.length}):
${itemsList}

Please check your dashboard to review and start processing the order.`;

        // Send messages (EXISTING - DO NOT CHANGE)
        const customerSent = await sendWhatsApp(customerPhone, customerMessage, customerName);

        // ✅ Send to MULTIPLE store owners (EXISTING - DO NOT CHANGE)
        const ownerResults = [];
        for (const ownerPhone of STORE_OWNERS) {
            const ownerSent = await sendWhatsApp(ownerPhone, ownerMessage, `Owner ${STORE_OWNERS.indexOf(ownerPhone) + 1}`);
            ownerResults.push(ownerSent);
        }

        const ownersSuccess = ownerResults.filter(result => result).length;
        return `customer:${customerSent}, owners:${ownersSuccess}/${STORE_OWNERS.length}`;

    } catch (error) {
        console.error('❌ Order processing error:', error);
        return 'error';
    }
}

// 🚚 HANDLE SHIPMENT UPDATE + GOOGLE SHEETS SYNC
async function handleShipmentUpdate(payload) {
    try {
        const { events, order } = payload;

        if (!order || !events) {
            return 'invalid_data';
        }

        const orderId = order.orderId;
        const customerPhone = getCustomerPhone(orderId);

        if (!customerPhone) {
            console.log(`❌ No phone found for order: ${orderId}`);
            return 'phone_not_found';
        }

        const customerName = customerPhones[orderId]?.name || 'there';
        const orderNumber = order.orderSerial;
        const shippingCompany = order.companyName || order.shippingRateName || 'shipping company';

        console.log('📦 Processing shipment events:', events);

        // Handle different shipment events
        for (const event of events) {
            switch (event) {
                case 'OrderShipmentPickedUp':
                    await handlePickupEvent(customerPhone, customerName, orderNumber, shippingCompany);
                    // 🆕 Update Google Sheets
                    await googleSheetsService.updateOrderStatus(orderId, 'SHIPPED');
                    break;

                case 'OrderShipmentDelivered':
                    await handleDeliveryEvent(customerPhone, customerName, orderNumber);
                    // 🆕 Update Google Sheets
                    await googleSheetsService.updateOrderStatus(orderId, 'DELIVERED');
                    break;

                case 'OrderShipmentIsOnTheWay':
                    // await handleInTransitEvent(customerPhone, customerName, orderNumber, shippingCompany);
                    break;

                default:
                    console.log(`📦 Other shipment event: ${event}`);
            }
        }

        return 'events_processed';

    } catch (error) {
        console.error('❌ Shipment processing error:', error);
        return 'error';
    }
}

// 🚚 PICKUP EVENT HANDLER (EXISTING - DO NOT CHANGE)
async function handlePickupEvent(customerPhone, customerName, orderNumber, shippingCompany) {
    const pickupMessage = `Hey ${customerName}! 💛  

Good news — your order #${orderNumber} has been picked up by ${shippingCompany} 🚚  

It's now on its way to you! 🕊  

Thank you for shopping with Glam&Glow ✨  
Enjoy your new items! 🕯`;

    return await sendWhatsApp(customerPhone, pickupMessage, customerName);
}

// 🎉 DELIVERY EVENT HANDLER (EXISTING - DO NOT CHANGE)
async function handleDeliveryEvent(customerPhone, customerName, orderNumber) {
    const deliveryMessage = `Hey ${customerName}! 💛  

We hope you're loving your new candles from Glam&Glow 🕯✨  
Your support truly means the world to us! 💫  

We'd love to invite you to join our Glam&Glow Family WhatsApp group 💬  
You'll get a permanent 10% OFF on all your future orders, plus early access to new launches and special offers! 🎁  

👉 Join here: https://chat.whatsapp.com/Gp4PBj6O2J8HTWgaHQ0S8i

Can't wait to see you there! 💛`;

    return await sendWhatsApp(customerPhone, deliveryMessage, customerName);
}

// ❌ HANDLE ORDER CANCELLATION - NOTIFY OWNERS + GOOGLE SHEETS
async function handleOrderCancel(order) {
    try {
        const orderId = order._id;
        const customerName = order.customer?.name || 'there';
        const rawPhone = order.shippingAddress?.phone;
        const customerPhone = formatPhone(rawPhone);

        if (!customerPhone) {
            return 'invalid_phone';
        }

        const orderNumber = order.orderSerial;

        // 🆕 Update Google Sheets
        await googleSheetsService.updateOrderStatus(orderId, 'CANCELED');

        // Customer cancellation message (EXISTING - DO NOT CHANGE)
        const cancelMessage = `Hey ${customerName} 💛

We wanted to let you know that your order #${orderNumber} has been canceled.

If this was a mistake or you have any questions, please reply to this message and we'll help you out!

Thank you for considering Glam&Glow ✨`;

        // Owner cancellation notification (EXISTING - DO NOT CHANGE)
        const ownerCancelMessage = `❌ Order Cancelled

Order #${orderNumber} has been canceled.

Customer: ${customerName}
Phone: ${rawPhone}

Reason: ${order.cancelReason || 'Not specified'}`;

        // Send to customer (EXISTING - DO NOT CHANGE)
        const customerSent = await sendWhatsApp(customerPhone, cancelMessage, customerName);

        // Send to all owners (EXISTING - DO NOT CHANGE)
        const ownerResults = [];
        for (const ownerPhone of STORE_OWNERS) {
            const ownerSent = await sendWhatsApp(ownerPhone, ownerCancelMessage, `Owner ${STORE_OWNERS.indexOf(ownerPhone) + 1}`);
            ownerResults.push(ownerSent);
        }

        // Remove from storage if canceled (EXISTING - DO NOT CHANGE)
        if (customerSent) {
            delete customerPhones[orderId];
            saveStorageData();
        }

        const ownersSuccess = ownerResults.filter(result => result).length;
        return `customer:${customerSent}, owners:${ownersSuccess}/${STORE_OWNERS.length}`;

    } catch (error) {
        console.error('❌ Cancel order error:', error);
        return 'error';
    }
}

// 💾 PHONE STORAGE FUNCTIONS (EXISTING - DO NOT CHANGE)
function storeCustomerPhone(orderId, phone, name) {
    customerPhones[orderId] = {
        phone: phone,
        name: name,
        storedAt: new Date().toISOString()
    };
    saveStorageData();
    console.log(`💾 Stored phone for ${name}`);
}

function getCustomerPhone(orderId) {
    return customerPhones[orderId]?.phone;
}

// 🔧 UTILITY FUNCTIONS (EXISTING - DO NOT CHANGE)
function formatPhone(rawPhone) {
    if (!rawPhone) return null;

    let cleaned = rawPhone.replace(/\D/g, '');

    if (cleaned.startsWith('0')) {
        cleaned = '2' + cleaned.substring(1);
    } else if (cleaned.startsWith('1') && cleaned.length === 10) {
        cleaned = '2' + cleaned;
    } else if (!cleaned.startsWith('2') && cleaned.length === 10) {
        cleaned = '2' + cleaned;
    }

    return cleaned.length >= 11 ? cleaned : null;
}

async function sendWhatsApp(phone, message, recipientName) {
    try {
        if (!client.info) {
            console.log('⚠️ WhatsApp client not ready');
            return false;
        }

        console.log(`📤 Sending to ${recipientName} (${phone})`);

        const numberId = await client.getNumberId(phone);
        if (!numberId) {
            console.log(`⚠️ Not on WhatsApp: ${phone}`);
            return false;
        }

        await client.sendMessage(
            numberId._serialized,
            message,
            {
                markSeen: false,
                sendSeen: false
            }
        );

        console.log(`✅ Sent to ${recipientName}`);
        return true;

    } catch (error) {
        console.error(`❌ Send failed to ${recipientName}:`, error);
        return false;
    }
}

// 🏥 HEALTH CHECK (ENHANCED WITH GOOGLE SHEETS STATUS)
app.get('/health', (req, res) => {
    res.json({
        status: 'OK',
        timestamp: new Date().toISOString(),
        whatsapp: client.info ? 'Connected' : 'Connecting',
        googleSheets: googleSheetsService.initialized ? 'Connected' : 'Not initialized',
        storage: {
            customerPhones: Object.keys(customerPhones).length,
            webhookLogs: Object.keys(webhookLogs).length
        },
        storeOwners: {
            count: STORE_OWNERS.length,
            numbers: STORE_OWNERS
        }
    });
});

// 👑 MANAGE STORE OWNERS ENDPOINTS (EXISTING - DO NOT CHANGE)
app.get('/owners', (req, res) => {
    res.json({
        owners: STORE_OWNERS.map((phone, index) => ({
            id: index + 1,
            phone: phone,
            status: 'active'
        }))
    });
});

// Add new owner (for future use) (EXISTING - DO NOT CHANGE)
app.post('/owners', (req, res) => {
    const { phone } = req.body;

    if (!phone) {
        return res.status(400).json({ error: 'Phone number required' });
    }

    const formattedPhone = formatPhone(phone);
    if (!formattedPhone) {
        return res.status(400).json({ error: 'Invalid phone number' });
    }

    if (STORE_OWNERS.includes(formattedPhone)) {
        return res.status(400).json({ error: 'Owner already exists' });
    }

    STORE_OWNERS.push(formattedPhone);

    res.json({
        message: 'Owner added successfully',
        totalOwners: STORE_OWNERS.length,
        newOwner: formattedPhone
    });
});

app.use((req, res) => {
    res.status(200).json({ status: 'OK' });
});

// 🚀 START SERVER
const PORT = 5000;
app.listen(PORT, () => {
    console.log(`
🕯 GLAM&GLOW WHATSAPP BOT STARTED (Enhanced with Google Sheets)
📍 Port: ${PORT}
📞 Webhook: POST http://localhost:${PORT}/webhooks/wuilt
❤️ Health: GET http://localhost:${PORT}/health
👑 Owners: GET http://localhost:${PORT}/owners
💾 Storage: ${Object.keys(customerPhones).length} customers
👥 Store Owners: ${STORE_OWNERS.length} owners
📊 Google Sheets: ${googleSheetsService.initialized ? 'Enabled' : 'Disabled'}
    `);

    // Display owner numbers (masked for security)
    STORE_OWNERS.forEach((phone, index) => {
        const maskedPhone = phone.substring(0, 6) + '****' + phone.substring(10);
        console.log(`   👑 Owner ${index + 1}: ${maskedPhone}`);
    });
});

// Graceful shutdown
process.on('SIGINT', async () => {
    console.log('🔄 Shutting down Glam&Glow bot...');
    saveStorageData();
    await client.destroy();
    process.exit(0);
});
