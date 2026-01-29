const express = require('express');
const cors = require('cors');
const db = require('./db');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Load SQL JSON file
// Load SQL JSON files
const sqlDir = path.join(__dirname, '../frontend/src/sql');
let sqlDefinitions = {};

try {
    const files = fs.readdirSync(sqlDir);
    files.forEach(file => {
        if (path.extname(file) === '.json') {
            const filePath = path.join(sqlDir, file);
            const data = fs.readFileSync(filePath, 'utf8');
            const queries = JSON.parse(data).queries;
            sqlDefinitions = { ...sqlDefinitions, ...queries };
            console.log(`Loaded SQL definitions from ${file}`);
        }
    });
    console.log('All SQL definitions loaded successfully');
} catch (err) {
    console.error('Error loading SQL definitions:', err);
}

// Initialize DB connection
db.initialize();

// Generic Query Endpoint
app.post('/api/query', async (req, res) => {
    const { queryId, params } = req.body;

    if (!queryId || !sqlDefinitions[queryId]) {
        return res.status(400).json({ error: 'Invalid or missing queryId' });
    }

    const queryDef = sqlDefinitions[queryId];
    const sql = queryDef.sql;

    // Convert named parameters if needed, or pass directly object if 'oracledb' supports bind by name
    // oracledb supports bind by name if object is passed.

    try {
        console.log(`Executing Query [${queryId}]:`, sql);
        console.log(`Params:`, params);

        const result = await db.execute(sql, params || {});
        res.json({ success: true, rows: result.rows });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// Health Check
app.get('/health', async (req, res) => {
    const isConnected = await db.checkConnection();
    res.json({
        status: isConnected ? 'ok' : 'error',
        db: isConnected ? 'connected' : 'disconnected'
    });
});

// Root Route
app.get('/', (req, res) => {
    res.send('Backend API Server is running. Access endpoints at /api/query');
});

process.on('SIGINT', async () => {
    console.log('Closing server...');
    await db.close();
    process.exit(0);
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
