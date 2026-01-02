const PDFToWordConverter = require('./converter');
const path = require('path');

async function testSinglePage() {
    const converter = new PDFToWordConverter();
    const input = path.resolve('d:\\تنسيق\\122.pdf');
    const output = path.resolve('d:\\تنسيق\\outputs\\test_122.docx');

    console.log("Testing Page 3 of 122.pdf...");
    await converter.convert(input, output);
    console.log("Done. Check outputs/test_122.docx");
}

testSinglePage().catch(console.error);
