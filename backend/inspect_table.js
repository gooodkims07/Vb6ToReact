const db = require('./db');

async function inspect() {
    try {
        await db.initialize();
        const sql = "SELECT IDNUMBER, NAME FROM TW_MIS_PMPA.TWBAS_PASS WHERE IDNUMBER = :sabun AND PASSWORD = :password";
        const binds = { sabun: '600018', password: 'devpass' };

        console.log('Testing Query:', sql);
        const result = await db.execute(sql, binds);
        console.log('Query Result:', JSON.stringify(result.rows, null, 2));
    } catch (err) {
        console.error('Inspection Error:', err);
    } finally {
        await db.close();
    }
}

inspect();
