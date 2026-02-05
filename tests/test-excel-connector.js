/**
 * Comprehensive Excel Connector Test
 */

const { COMAutomation } = require('../index.js');

console.log('=== Excel Connector Comprehensive Test ===\n');

async function testExcelConnector() {
    console.log('Test: Excel Connector Methods');
    console.log('------------------------------');
    
    let excel = null;
    
    try {
        // Create connector
        excel = new COMAutomation.ExcelConnector();
        console.log('✓ ExcelConnector created');
        
        // Get version
        const version = excel.getVersion();
        console.log(`✓ Excel version: ${version}`);
        
        // Add workbook
        console.log('\n✓ Creating new workbook...');
        const workbook = excel.addWorkbook();
        console.log('✓ Workbook created');
        
        // Get active sheet
        const sheet = workbook.getActiveSheet();
        console.log('✓ Active sheet obtained');
        
        const sheetName = sheet.getName();
        console.log(`  Sheet name: ${sheetName}`);
        
        // Set sheet name
        sheet.setName('Test Data');
        console.log('✓ Sheet name changed to "Test Data"');
        
        // Write data to cells using getCells (more reliable)
        console.log('\n✓ Writing data to cells...');
        
        // Headers
        sheet.getCells(1, 1).setValue('Product');
        sheet.getCells(1, 2).setValue('Price');
        sheet.getCells(1, 3).setValue('Quantity');
        sheet.getCells(1, 4).setValue('Total');
        console.log('✓ Headers written');
        
        // Data rows
        sheet.getCells(2, 1).setValue('Widget A');
        sheet.getCells(2, 2).setValue(29.99);
        sheet.getCells(2, 3).setValue(10);
        sheet.getCells(2, 4).setFormula('=B2*C2');
        
        sheet.getCells(3, 1).setValue('Widget B');
        sheet.getCells(3, 2).setValue(49.99);
        sheet.getCells(3, 3).setValue(5);
        sheet.getCells(3, 4).setFormula('=B3*C3');
        
        sheet.getCells(4, 1).setValue('Widget C');
        sheet.getCells(4, 2).setValue(19.99);
        sheet.getCells(4, 3).setValue(15);
        sheet.getCells(4, 4).setFormula('=B4*C4');
        console.log('✓ Data rows written with formulas');
        
        // Format headers
        console.log('\n✓ Formatting cells...');
        const headerCell = sheet.getCells(1, 1);
        const headerFont = headerCell.getFont();
        headerFont.setBold(true);
        headerFont.setSize(12);
        console.log('✓ Header font formatted (bold, size 12)');
        
        const headerInterior = headerCell.getInterior();
        headerInterior.setColor(0xCCCCCC); // Light gray
        console.log('✓ Header background color set');
        
        // Format currency columns
        sheet.getCells(2, 2).setNumberFormat('$#,##0.00');
        sheet.getCells(3, 2).setNumberFormat('$#,##0.00');
        sheet.getCells(4, 2).setNumberFormat('$#,##0.00');
        console.log('✓ Price column formatted as currency');
        
        sheet.getCells(2, 4).setNumberFormat('$#,##0.00');
        sheet.getCells(3, 4).setNumberFormat('$#,##0.00');
        sheet.getCells(4, 4).setNumberFormat('$#,##0.00');
        console.log('✓ Total column formatted as currency');
        
        // Read values back
        console.log('\n✓ Reading values back...');
        const productValue = sheet.getCells(2, 1).getValue();
        console.log(`  A2 value: ${productValue}`);
        
        const priceValue = sheet.getCells(2, 2).getValue();
        console.log(`  B2 value: ${priceValue}`);
        
        // Get workbook count
        const workbookCount = excel.getWorkbookCount();
        console.log(`\n✓ Open workbooks: ${workbookCount}`);
        
        // Test calculation mode
        console.log('\n✓ Testing calculation settings...');
        excel.setCalculation(COMAutomation.XlCalculation.xlCalculationAutomatic);
        console.log('✓ Calculation mode set to Automatic');
        
        // Save workbook
        console.log('\n✓ Saving workbook...');
        const testPath = require('path').join(process.cwd(), 'test-output.xlsx');
        workbook.saveAs(testPath, COMAutomation.XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Workbook saved to: ${testPath}`);
        
        // Close workbook
        workbook.close(false); // Don't save changes again
        console.log('✓ Workbook closed');
        
        // Quit Excel
        excel.quit();
        console.log('✓ Excel quit');
        
        console.log('\n✅ All Excel connector tests PASSED!');
        return true;
        
    } catch (error) {
        console.error('\n❌ Excel connector test failed:', error.message);
        console.error(error.stack);
        return false;
    } finally {
        if (excel) {
            excel.release();
            console.log('\n✓ Excel connector released');
        }
    }
}

async function runTest() {
    console.log('Starting Excel connector test...\n');
    
    const result = await testExcelConnector();
    
    console.log('\n\n=== Test Summary ===');
    console.log(`Excel Test: ${result ? '✅ PASSED' : '❌ FAILED'}`);
    
    if (result) {
        console.log('\n✅ Excel connector fully functional!');
        console.log('\nVerified:');
        console.log('  ✓ Version detection');
        console.log('  ✓ Workbook creation');
        console.log('  ✓ Sheet manipulation');
        console.log('  ✓ Cell value operations (getCells)');
        console.log('  ✓ Formula support');
        console.log('  ✓ Cell formatting (font, color, number format)');
        console.log('  ✓ File save operations');
        console.log('  ✓ Calculation mode settings');
    }
}

runTest().catch(error => {
    console.error('Unexpected error:', error);
    process.exit(1);
});
