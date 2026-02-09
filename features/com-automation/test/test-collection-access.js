// Test collection access with DISPID_VALUE
const { COMObject } = require('../../../build/Release/Automation.node');

console.log('Testing Excel collection access with DISPID_VALUE...\n');

const excel = new COMObject('Excel.Application');
excel.setProperty('Visible', true);
excel.setProperty('DisplayAlerts', false);

const workbooks = excel.getProperty('Workbooks');
const workbook = workbooks.invoke('Add');

const worksheets = workbook.getProperty('Worksheets');
console.log('Worksheets collection obtained');

const count = worksheets.getProperty('Count');
console.log('Count:', count);

// Try using DISPID_VALUE (0) for default property
try {
    console.log('\nTrying DISPID_VALUE approach...');
    // In COM, default properties can be accessed via DISPID_VALUE
    // We need to use invokeWithDispId or similar
    
    // Let's try getting the first item using different approaches
    const sheet = worksheets.invokeWithDispId(0, 1); // DISPID_VALUE = 0
    console.log('Success with DISPID_VALUE!');
    console.log('Sheet name:', sheet.getProperty('Name'));
} catch (e) {
    console.log('Failed with DISPID_VALUE:', e.message);
}

// Try the _NewEnum approach
try {
    console.log('\nTrying enumeration...');
    const enumVariant = worksheets.invoke('_NewEnum');
    console.log('Got enumerator:', enumVariant);
} catch (e) {
    console.log('Failed to get enumerator:', e.message);
}

// Try direct property access with brackets
try {
    console.log('\nTrying bracket notation simulation...');
    // Collections in Excel are 1-based
    // Try calling the default member
    const sheet = worksheets.invoke(1);
    console.log('Success with invoke(1)!');
    console.log('Sheet name:', sheet.getProperty('Name'));
} catch (e) {
    console.log('Failed with invoke(1):', e.message);
}

workbook.invoke('Close', false);
excel.invoke('Quit');
excel.release();
