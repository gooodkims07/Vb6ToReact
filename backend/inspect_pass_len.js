const db = require('./db');

async function checkPass() {
    try {
        await db.initialize();
        const sql = "SELECT PASSWORD FROM TW_MIS_PMPA.TWBAS_PASS WHERE IDNUMBER = :sabun";
        const binds = { sabun: '600018' };

        const result = await db.execute(sql, binds);
        console.log('Use ID: 600018');
        result.rows.forEach(r => {
            console.log(`Password: '${r.PASSWORD}', Length: ${r.PASSWORD.length}`);
        });
    } catch (err) {
        console.error('Error:', err);
    } finally {
        await db.close();
    }
}

checkPass();
