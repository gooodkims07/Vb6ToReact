const oracledb = require('oracledb');
const path = require('path');
const fs = require('fs');
require('dotenv').config({ path: path.join(__dirname, '.env') });

// Enable Thick mode (required for older DB versions)
try {
    oracledb.initOracleClient({ libDir: 'C:\\oracle\\client_x64\\product\\12.1.0\\client_1\\bin' });
} catch (err) {
    console.error('Oracle Client init error:', err);
}

const dbConfig = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    connectString: `${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_SID}`
};

async function initialize() {
    try {
        // Create a connection pool which will later be accessed by the execution logic
        await oracledb.createPool({
            user: dbConfig.user,
            password: dbConfig.password,
            connectString: dbConfig.connectString,
            poolMin: 2,
            poolMax: 10,
            poolIncrement: 1
        });
        console.log('Oracle Database pool created');
    } catch (err) {
        console.error('init() error: ' + err.message);
    }
}

async function close() {
    try {
        await oracledb.getPool().close(10);
        console.log('Pool closed');
    } catch (err) {
        console.error('close() error: ' + err.message);
    }
}

async function execute(sql, binds = [], options = {}) {
    let connection;
    try {
        connection = await oracledb.getConnection();

        const start = Date.now();
        const logMsg = `\n[${new Date().toISOString()}] SQL: ${sql}\nParams: ${JSON.stringify(binds)}\n`;
        fs.appendFileSync(path.join(__dirname, 'backend_debug.log'), logMsg);

        console.log('\n================ DATA QUERY LOG ================');
        console.log(`[${new Date().toISOString()}] Executing SQL:`);
        console.log(sql);
        console.log('------------------------------------------------');
        console.log('Bind Parameters:', JSON.stringify(binds));
        console.log('================================================\n');

        const result = await connection.execute(sql, binds, {
            outFormat: oracledb.OUT_FORMAT_OBJECT, // Return result as JSON objects
            autoCommit: true,
            ...options
        });

        const duration = Date.now() - start;
        console.log(`Query completed in ${duration}ms. Rows: ${result.rows ? result.rows.length : 0}\n`);

        return result;
    } catch (err) {
        console.error('Execute error:', err);
        throw err;
    } finally {
        if (connection) {
            try {
                await connection.close();
            } catch (err) {
                console.error(err);
            }
        }
    }
}

async function checkConnection() {
    let connection;
    try {
        connection = await oracledb.getConnection();
        await connection.ping();
        return true;
    } catch (err) {
        console.error("Check Connection Error:", err);
        return false;
    } finally {
        if (connection) {
            try {
                await connection.close();
            } catch (err) {
                console.error(err);
            }
        }
    }
}

module.exports = {
    initialize,
    close,
    execute,
    checkConnection
};
