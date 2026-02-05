/**
 * Test saveAs method specifically
 */

const { COMAutomation } = require('../index.js');
const path = require('path');

async function testSaveAs() {
    const excel = new COMAutomation.ExcelConnector();
    
    try {
        excel.setVisible(true);
        
        const workbook = excel.addWorkbook();
        console.log('✓ Workbook created');
        
        const sheet = workbook.getActiveSheet();
        sheet.setName('Test');
        console.log('✓ Sheet renamed');
        
        // Try save() first
        try {
            workbook.save();
            console.log('✓ save() works');
        } catch (e) {
            console.log(`✗ save() failed: ${e.message}`);
        }
        
        // Try saveAs()
        try {
            const testPath = path.join(process.cwd(), 'test-saveas.xlsx');
            console.log(`Attempting saveAs to: ${testPath}`);
            workbook.saveAs(testPath, COMAutomation.XlFileFormat.xlOpenXMLWorkbook);
            console.log('✓ saveAs() works');
        } catch (e) {
            console.log(`✗ saveAs() failed: ${e.message}`);
        }
        
        workbook.close(false);
        
    } finally {
        excel.quit();
        excel.release();
    }
}

testSaveAs().catch(console.error);
