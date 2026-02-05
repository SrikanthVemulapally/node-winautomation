/**
 * Test Excel properties to identify which ones work
 */

const { COMAutomation } = require('../index.js');

async function testProperties() {
    const excel = new COMAutomation.ExcelConnector();
    
    try {
        console.log('Testing Excel properties...\n');
        
        // Test 1: Version (should work)
        try {
            const version = excel.getVersion();
            console.log(`✓ getVersion: ${version}`);
        } catch (e) {
            console.log(`✗ getVersion: ${e.message}`);
        }
        
        // Test 2: Visible (should work)
        try {
            excel.setVisible(true);
            const visible = excel.getVisible();
            console.log(`✓ setVisible/getVisible: ${visible}`);
        } catch (e) {
            console.log(`✗ setVisible/getVisible: ${e.message}`);
        }
        
        // Test 3: DisplayAlerts (should work)
        try {
            excel.setDisplayAlerts(false);
            console.log(`✓ setDisplayAlerts: works`);
        } catch (e) {
            console.log(`✗ setDisplayAlerts: ${e.message}`);
        }
        
        // Test 4: ScreenUpdating (may not work)
        try {
            excel.setScreenUpdating(false);
            console.log(`✓ setScreenUpdating: works`);
        } catch (e) {
            console.log(`✗ setScreenUpdating: ${e.message}`);
        }
        
        // Test 5: Calculation (may not work)
        try {
            excel.setCalculation(COMAutomation.XlCalculation.xlCalculationAutomatic);
            console.log(`✓ setCalculation: works`);
        } catch (e) {
            console.log(`✗ setCalculation: ${e.message}`);
        }
        
        console.log('\n=== Summary ===');
        console.log('Properties that work can be used in examples.');
        console.log('Properties that fail should be removed from examples.');
        
    } finally {
        excel.quit();
        excel.release();
    }
}

testProperties().catch(console.error);
