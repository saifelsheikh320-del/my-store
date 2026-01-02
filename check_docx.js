const fs = require('fs');

function checkFile() {
    const stats = fs.statSync('d:\\تنسيق\\outputs\\122.docx');
    console.log(`File Size: ${stats.size} bytes`);
    if (stats.size < 10000) {
        console.log("Verdict: The file is likely too small to contain 10 pages of extracted text.");
    } else {
        console.log("Verdict: The file has some content.");
    }
}

checkFile();
