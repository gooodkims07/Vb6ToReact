const db = require('./db');

async function checkPass() {
    try {
        await db.initialize();
        const sql = "SELECT PASSWORD, DUMP(PASSWORD) as DUMP_VAL FROM TW_MIS_PMPA.TWBAS_PASS WHERE IDNUMBER = :sabun";
        const binds = { sabun: '600018' };

        const result = await db.execute(sql, binds);
        console.log('Result:', JSON.stringify(result.rows, null, 2));
    } catch (err) {
        console.error('Error:', err);
    } finally {
        await db.close();
    }
}

checkPass();
