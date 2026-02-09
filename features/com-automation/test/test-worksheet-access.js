// Debug worksheet access issue
const { ExcelConnector } = require('../lib/connectors/ExcelConnector');

console.log('Testing worksheet access...\n');

const excel = new ExcelConnector();
excel.setVisible(true);
excel.setDisplayAlerts(false);

const workbook = excel.addWorkbook();
console.log('Workbook created:', workbook.getName());

const worksheets = workbook.getWorksheets();
console.log('Worksheets collection:', worksheets);

const count = worksheets.getProperty('Count');
console.log('Worksheet count:', count);

// Try different ways to access worksheet
console.log('\nTrying different access methods:');

try {
    console.log('1. Using invoke("Item", 1)...');
    const sheet1 = worksheets.invoke('Item', 1);
    console.log('   Success! Sheet:', sheet1);
} catch (e) {
    console.log('   Failed:', e.message);
}

try {
    console.log('2. Using getProperty("Item", 1)...');
    const sheet2 = worksheets.getProperty('Item', 1);
    console.log('   Success! Sheet:', sheet2);
} catch (e) {
    console.log('   Failed:', e.message);
}

try {
    console.log('3. Using invoke("_Default", 1)...');
    const sheet3 = worksheets.invoke('_Default', 1);
    console.log('   Success! Sheet:', sheet3);
} catch (e) {
    console.log('   Failed:', e.message);
}

try {
    console.log('4. Using getProperty with index directly...');
    const sheet4 = worksheets.getProperty(1);
    console.log('   Success! Sheet:', sheet4);
} catch (e) {
    console.log('   Failed:', e.message);
}

workbook.close(false);
excel.quit();
excel.release();
