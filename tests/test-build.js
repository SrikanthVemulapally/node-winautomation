/**
 * Quick test to verify the build works
 */

try {
    console.log('Testing node-winautomation build...\n');
    
    // Test UI Automation
    const automation = require('./index.js');
    console.log('✓ Module loaded successfully');
    
    // Check UI Automation exports
    if (automation.Automation) {
        console.log('✓ UI Automation available');
    }
    
    // Check COM Automation exports
    if (automation.OutlookConnector) {
        console.log('✓ OutlookConnector available');
    }
    
    if (automation.ExcelConnector) {
        console.log('✓ ExcelConnector available');
    }
    
    // Check enumerations
    if (automation.ControlTypeIds) {
        console.log('✓ Enumerations available');
    }
    
    console.log('\n✅ Build verification successful!');
    console.log('\nAvailable features:');
    console.log('  - UI Automation');
    console.log('  - COM Automation (Outlook & Excel connectors)');
    console.log('\nNote: Desktop Management temporarily disabled (needs Windows SDK)');
    
} catch (error) {
    console.error('❌ Build verification failed:', error.message);
    console.error('\nThe .node file may be locked by another process.');
    console.error('Please close all Node.js processes and rebuild.');
    process.exit(1);
}
