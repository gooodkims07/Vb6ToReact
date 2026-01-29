const fs = require('fs');
const iconv = require('iconv-lite');
const path = require('path');

const vbFilePath = path.join(__dirname, '../vb6/FrmAttend1.frm');
const contentBuffer = fs.readFileSync(vbFilePath);
const content = iconv.decode(contentBuffer, 'EUC-KR'); // Common encoding for Korean VB6

const lines = content.split(/\r?\n/);
let extractedParams = [];
let currentSql = "";
let inSqlBlock = false;

console.log("--- EXTRACTION_START ---");

lines.forEach((line, index) => {
    let trimLine = line.trim();

    // Simple state machine to capture SQL blocks
    // Matches: strSql = "SELECT ..." or SQL = "..."
    if (trimLine.match(/^(strSql|SQL)\s*=\s*"/i)) {
        if (inSqlBlock && currentSql.length > 5) {
            console.log("SQL_BLOCK:");
            console.log(currentSql);
        }
        currentSql = "";
        let match = trimLine.match(/=\s*"(.*)"/);
        if (match) currentSql = match[1];
        inSqlBlock = true;
    }
    // Matches: strSql = strSql & "..."
    else if (inSqlBlock && trimLine.match(/^(strSql|SQL)\s*=\s*(strSql|SQL)\s*&\s*"/i)) {
        let match = trimLine.match(/&\s*"(.*)"/);
        if (match) currentSql += " " + match[1];
    }
    // Assignment to something else ends the block
    else if (inSqlBlock && trimLine.match(/^\w+\s*=/)) {
        if (currentSql.length > 5) {
            console.log("SQL_BLOCK:");
            console.log(currentSql);
        }
        inSqlBlock = false;
        currentSql = "";
    }
    // End Sub/Function ends block
    else if (trimLine.match(/^End (Sub|Function)/i)) {
        if (inSqlBlock && currentSql.length > 5) {
            console.log("SQL_BLOCK:");
            console.log(currentSql);
        }
        inSqlBlock = false;
        currentSql = "";
    }
});

if (inSqlBlock && currentSql.length > 5) {
    console.log("SQL_BLOCK:");
    console.log(currentSql);
}

console.log("--- EXTRACTION_END ---");
