const db = require('./db');

async function testFetch() {
    try {
        await db.initialize();

        // 1. Analyze Data
        console.log("--- Analyzing Departments ---");
        const analyzeSql = "SELECT DEPT, COUNT(*) as CNT FROM TW_MIS_ADM.TWINSA_MASTER GROUP BY DEPT ORDER BY CNT DESC";
        const analyzeResult = await db.execute(analyzeSql, {}, { maxRows: 5 });
        console.log("Top Depts:", analyzeResult.rows);

        if (analyzeResult.rows.length > 0) {
            const targetDept = analyzeResult.rows[0].DEPT;
            console.log(`\n--- Target Dept: '${targetDept}' (Length: ${targetDept.length}) ---`);

            // 4. Reproduction Test (ORA-01036)
            console.log("--- Reproduction Test ---");
            // SQL exactly as in JSON: :2 first (in Join), :1 second (in WHERE with RPAD)
            const reproSql = "SELECT A.SABUN, A.NAMEK FROM TW_MIS_ADM.TWINSA_MASTER A LEFT JOIN TW_MIS_ADM.TWINSA_WORKTIME C ON A.SABUN = C.SABUN AND C.WORKDATE = :2 WHERE A.DEPT = RPAD(:1, 6, ' ') ORDER BY A.SABUN";

            // Params exactly as in JSX (after swap): [Date, Dept]
            const params = ['2023-10-11', targetDept];

            console.log("Executing with SQL:", reproSql);
            console.log("Params:", params);

            const reproResult = await db.execute(reproSql, params);
            console.log(`Reproduction Count: ${reproResult.rows.length}`);
        } else {
            console.log("No departments found.");
        }

    } catch (err) {
        console.error("Test Failed:", err);
    } finally {
        await db.close();
    }
}

testFetch();
