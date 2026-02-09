/**
 * Debug writeRange/readRange issue
 */

const { ExcelConnector } = require('../index.js');

const excel = new ExcelConnector();

try {
    excel.setVisible(false);
    excel.setDisplayAlerts(false);
    
    const workbook = excel.addWorkbook();
    const sheet = workbook.getActiveSheet();
    
    console.log('Testing writeRange/readRange...');
    
    const data = [
        ['Name', 'Age', 'City'],
        ['John', 30, 'NYC'],
        ['Jane', 25, 'LA']
    ];
    
    console.log('Writing data:', data);
    sheet.writeRange('A1:C3', data);
    
    console.log('Reading data back...');
    const readData = sheet.readRange('A1:C3');
    console.log('Read data:', readData);
    console.log('Type:', typeof readData);
    console.log('Is null:', readData === null);
    console.log('Is undefined:', readData === undefined);
    
    // Try reading individual cells
    console.log('\nReading individual cells:');
    console.log('A1:', sheet.readCell(1, 1));
    console.log('B1:', sheet.readCell(1, 2));
    console.log('C1:', sheet.readCell(1, 3));
    
    workbook.close(false);
} finally {
    excel.quit();
    excel.release();
}
