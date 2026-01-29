const db = require('./db');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });
const http = require('http');

async function test() {
    console.log("1. Testing DB Connection Logic directly...");
    await db.initialize();
    const isConnected = await db.checkConnection();
    console.log("Direct DB Check Result:", isConnected ? "SUCCESS" : "FAILED");

    console.log("2. Testing Health Endpoint...");
    const req = http.get('http://localhost:3000/health', (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', () => {
            console.log("Health Endpoint Result:", data);
            process.exit(0);
        });
    });

    req.on('error', (e) => {
        console.error("Health Endpoint Failed:", e.message);
        process.exit(1);
    });
}

test();
