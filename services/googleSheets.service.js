// services/googleSheets.service.js
const { google } = require('googleapis');
const path = require('path');
require('dotenv').config({ path: './key/.env' });


class GoogleSheetsService {
    constructor() {
        this.sheets = null;
        this.spreadsheetId = process.env.GOOGLE_SHEET_ID;
        this.sheetName = 'Orders';
        this.initialized = false;
    }

    async initialize() {
        try {
            // Load service account credentials
            const auth = new google.auth.GoogleAuth({
                keyFile: path.join(__dirname, '../service/google-credentials.json'),
                scopes: ['https://www.googleapis.com/auth/spreadsheets'],
            });

            const client = await auth.getClient();
            this.sheets = google.sheets({ version: 'v4', auth: client });
            this.initialized = true;
            
            console.log('✅ Google Sheets initialized successfully');
            
            // Ensure headers exist
            await this.ensureHeaders();
            
            return true;
        } catch (error) {
            console.error('❌ Google Sheets initialization failed:', error.message);
            this.initialized = false;
            return false;
        }
    }

    async ensureHeaders() {
        if (!this.initialized) return;

        try {
            const headers = [
                'order_id',
                'order_number',
                'customer_name',
                'phone',
                'total_amount',
                'items_summary',
                'status',
                'last_customer_message',
                'confirmed_at',
                'created_at'
            ];

            // Check if sheet exists
            const response = await this.sheets.spreadsheets.get({
                spreadsheetId: this.spreadsheetId,
            });

            const sheetExists = response.data.sheets.some(
                sheet => sheet.properties.title === this.sheetName
            );

            if (!sheetExists) {
                // Create sheet
                await this.sheets.spreadsheets.batchUpdate({
                    spreadsheetId: this.spreadsheetId,
                    resource: {
                        requests: [{
                            addSheet: {
                                properties: {
                                    title: this.sheetName
                                }
                            }
                        }]
                    }
                });
                console.log(`📊 Created sheet: ${this.sheetName}`);
            }

            // Get existing headers
            const headerResponse = await this.sheets.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!A1:J1`,
            });

            const existingHeaders = headerResponse.data.values?.[0] || [];

            // Add headers if empty
            if (existingHeaders.length === 0) {
                await this.sheets.spreadsheets.values.update({
                    spreadsheetId: this.spreadsheetId,
                    range: `${this.sheetName}!A1:J1`,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [headers]
                    }
                });
                console.log('📋 Headers added to Google Sheet');
            }

        } catch (error) {
            console.error('❌ Error ensuring headers:', error.message);
        }
    }

    async addOrder(order) {
        if (!this.initialized) {
            console.log('⚠️ Google Sheets not initialized - skipping addOrder');
            return false;
        }

        try {
            const orderId = order._id;
            const orderNumber = order.orderSerial;
            const customerName = order.customer?.name || '';
            const phone = order.shippingAddress?.phone || order.customer?.phone || '';
            const totalAmount = order.totalPrice?.amount || 0;
            
            // Create items summary
            const itemsSummary = order.items?.map(item => {
                const sizeOption = item.options?.find(opt => opt.name === 'Size');
                const size = sizeOption ? ` (${sizeOption.value})` : '';
                return `${item.title}${size} x${item.quantity}`;
            }).join(', ') || '';

            const status = 'PENDING_CONFIRMATION';
            const createdAt = new Date().toISOString();

            const row = [
                orderId,
                orderNumber,
                customerName,
                phone,
                totalAmount,
                itemsSummary,
                status,
                '', // last_customer_message
                '', // confirmed_at
                createdAt
            ];

            await this.sheets.spreadsheets.values.append({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!A:J`,
                valueInputOption: 'RAW',
                insertDataOption: 'INSERT_ROWS',
                resource: {
                    values: [row]
                }
            });

            console.log(`📊 Added order ${orderNumber} to Google Sheet`);
            return true;

        } catch (error) {
            console.error('❌ Error adding order to sheet:', error.message);
            return false;
        }
    }

    async updateOrderStatus(orderId, status, additionalData = {}) {
        if (!this.initialized) {
            console.log('⚠️ Google Sheets not initialized - skipping updateOrderStatus');
            return false;
        }

        try {
            const rowIndex = await this.findRowByOrderId(orderId);
            
            if (rowIndex === -1) {
                console.log(`⚠️ Order ${orderId} not found in sheet`);
                return false;
            }

            // Get current row data
            const currentRow = await this.sheets.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!A${rowIndex}:J${rowIndex}`,
            });

            const row = currentRow.data.values[0];

            // Update status (column G = index 6)
            row[6] = status;

            // Update confirmed_at if status is CONFIRMED (column I = index 8)
            if (status === 'CONFIRMED' && !row[8]) {
                row[8] = new Date().toISOString();
            }

            // Update additional data if provided
            if (additionalData.confirmedAt) {
                row[8] = additionalData.confirmedAt;
            }

            await this.sheets.spreadsheets.values.update({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!A${rowIndex}:J${rowIndex}`,
                valueInputOption: 'RAW',
                resource: {
                    values: [row]
                }
            });

            console.log(`📊 Updated order ${orderId} status to ${status}`);
            return true;

        } catch (error) {
            console.error('❌ Error updating order status:', error.message);
            return false;
        }
    }

    async updateLastMessage(orderId, message) {
        if (!this.initialized) {
            console.log('⚠️ Google Sheets not initialized - skipping updateLastMessage');
            return false;
        }

        try {
            const rowIndex = await this.findRowByOrderId(orderId);
            
            if (rowIndex === -1) {
                console.log(`⚠️ Order ${orderId} not found in sheet`);
                return false;
            }

            // Update last_customer_message (column H = index 7)
            await this.sheets.spreadsheets.values.update({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!H${rowIndex}`,
                valueInputOption: 'RAW',
                resource: {
                    values: [[message]]
                }
            });

            console.log(`📊 Updated last message for order ${orderId}`);
            return true;

        } catch (error) {
            console.error('❌ Error updating last message:', error.message);
            return false;
        }
    }

    async getPendingOrderByPhone(phone) {
        if (!this.initialized) {
            console.log('⚠️ Google Sheets not initialized - skipping getPendingOrderByPhone');
            return null;
        }

        try {
            // Get all data
            const response = await this.sheets.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!A2:J`, // Skip header row
            });

            const rows = response.data.values || [];

            // Clean phone number for comparison
            const cleanPhone = phone.replace(/\D/g, '');

            // Find latest pending order matching this phone
            let latestOrder = null;
            let latestDate = null;

            for (let i = rows.length - 1; i >= 0; i--) {
                const row = rows[i];
                const rowPhone = (row[3] || '').replace(/\D/g, '');
                const status = row[6];
                const createdAt = row[9];

                // Match phone and pending status
                if (rowPhone.includes(cleanPhone) || cleanPhone.includes(rowPhone)) {
                    if (status === 'PENDING_CONFIRMATION') {
                        if (!latestDate || createdAt > latestDate) {
                            latestOrder = {
                                orderId: row[0],
                                orderNumber: row[1],
                                customerName: row[2],
                                phone: row[3],
                                totalAmount: row[4],
                                itemsSummary: row[5],
                                status: row[6],
                                lastCustomerMessage: row[7] || '',
                                confirmedAt: row[8] || '',
                                createdAt: row[9]
                            };
                            latestDate = createdAt;
                        }
                    }
                }
            }

            if (latestOrder) {
                console.log(`📊 Found pending order ${latestOrder.orderNumber} for phone ${phone}`);
            }

            return latestOrder;

        } catch (error) {
            console.error('❌ Error getting pending order:', error.message);
            return null;
        }
    }

    async findRowByOrderId(orderId) {
        try {
            const response = await this.sheets.spreadsheets.values.get({
                spreadsheetId: this.spreadsheetId,
                range: `${this.sheetName}!A:A`,
            });

            const rows = response.data.values || [];

            for (let i = 0; i < rows.length; i++) {
                if (rows[i][0] === orderId) {
                    return i + 1; // +1 because sheets are 1-indexed
                }
            }

            return -1;

        } catch (error) {
            console.error('❌ Error finding row:', error.message);
            return -1;
        }
    }
}

// Export singleton instance
module.exports = new GoogleSheetsService();
