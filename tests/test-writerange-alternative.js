/**
 * Test alternative writeRange approach
 */

const { ExcelConnector } = require('../index.js');

const excel = new ExcelConnector();

try {
    excel.setVisible(true);
    excel.setDisplayAlerts(false);
    
    const workbook = excel.addWorkbook();
    const sheet = workbook.getActiveSheet();
    
    console.log('Testing alternative writeRange approach...');
    
    // Write data cell by cell first
    sheet.writeCell(1, 1, 'Name');
    sheet.writeCell(1, 2, 'Age');
    sheet.writeCell(1, 3, 'City');
    sheet.writeCell(2, 1, 'John');
    sheet.writeCell(2, 2, 30);
    sheet.writeCell(2, 3, 'NYC');
    
    console.log('Data written cell by cell');
    
    // Now try to read the range
    console.log('Reading range A1:C2...');
    const readData = sheet.readRange('A1:C2');
    console.log('Read data:', readData);
    console.log('Type:', typeof readData);
    
    // Wait for user to see
    setTimeout(() => {
        workbook.close(false);
        excel.quit();
        excel.release();
    }, 3000);
    
} catch (e) {
    console.error('Error:', e.message);
    excel.quit();
    excel.release();
}
