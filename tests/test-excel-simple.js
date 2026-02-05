/**
 * Simple Excel Connector Test
 * Tests basic Excel operations
 */

const { COMAutomation } = require('../index.js');

console.log('=== Excel Connector Test ===\n');

async function testExcel() {
    let excel = null;
    
    try {
        // Create connector
        excel = new COMAutomation.ExcelConnector();
        console.log('✓ ExcelConnector created');
        
        // Get version
        const version = excel.getVersion();
        console.log(`✓ Excel version: ${version}`);
        
        // Set visible
        excel.setVisible(true);
        console.log('✓ Excel set to visible');
        
        // Add workbook
        const workbook = excel.addWorkbook();
        console.log('✓ Workbook created');
        
        // Get active sheet
        const sheet = workbook.getActiveSheet();
        console.log('✓ Active sheet obtained');
        
        const sheetName = sheet.getName();
        console.log(`  Sheet name: ${sheetName}`);
        
        // Set sheet name
        sheet.setName('Test Data');
        console.log('✓ Sheet name changed');
        
        // Test calculation mode
        excel.setCalculation(COMAutomation.XlCalculation.xlCalculationAutomatic);
        console.log('✓ Calculation mode set');
        
        // Save workbook
        const testPath = require('path').join(process.cwd(), 'test-excel-output.xlsx');
        workbook.saveAs(testPath, COMAutomation.XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Workbook saved to: ${testPath}`);
        
        // Close workbook
        workbook.close(false);
        console.log('✓ Workbook closed');
        
        // Quit Excel
        excel.quit();
        console.log('✓ Excel quit');
        
        console.log('\n✅ Excel connector test PASSED!');
        return true;
        
    } catch (error) {
        console.error('\n❌ Test failed:', error.message);
        return false;
    } finally {
        if (excel) {
            excel.release();
            console.log('\n✓ Excel connector released');
        }
    }
}

testExcel().then(result => {
    console.log('\n=== Test Summary ===');
    console.log(`Excel Test: ${result ? '✅ PASSED' : '❌ FAILED'}`);
    
    if (result) {
        console.log('\nVerified:');
        console.log('  ✓ Version detection');
        console.log('  ✓ Workbook creation');
        console.log('  ✓ Sheet access and naming');
        console.log('  ✓ Calculation mode settings');
        console.log('  ✓ File save operations');
        console.log('  ✓ Workbook close and Excel quit');
        console.log('\n⚠ Note: Cell operations (Range/Cells) require parameterized property support');
    }
}).catch(error => {
    console.error('Unexpected error:', error);
    process.exit(1);
});
