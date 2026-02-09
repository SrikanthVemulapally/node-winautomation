/**
 * Comprehensive test suite for ExcelConnector.js
 * Tests all classes and methods
 */

const path = require('path');
const fs = require('fs');
const os = require('os');
const {
    ExcelConnector,
    XlFileFormat,
    XlChartType,
    XlCalculation
} = require('../lib/connectors/ExcelConnector');

// Test output directory
const testDir = path.join(os.tmpdir(), 'excel-connector-tests');
if (!fs.existsSync(testDir)) {
    fs.mkdirSync(testDir, { recursive: true });
}

console.log('='.repeat(70));
console.log('COMPREHENSIVE EXCEL CONNECTOR TEST SUITE');
console.log('='.repeat(70));
console.log(`Test directory: ${testDir}\n`);

let testsPassed = 0;
let testsFailed = 0;
let excel = null;
let workbook = null;

function test(name, fn) {
    try {
        console.log(`\n[TEST] ${name}`);
        fn();
        console.log(`  ✓ PASSED`);
        testsPassed++;
    } catch (error) {
        console.log(`  ✗ FAILED: ${error.message}`);
        testsFailed++;
    }
}

function assert(condition, message) {
    if (!condition) {
        throw new Error(message || 'Assertion failed');
    }
}

function assertEquals(actual, expected, message) {
    if (actual !== expected) {
        throw new Error(message || `Expected ${expected}, got ${actual}`);
    }
}

function assertNotNull(value, message) {
    if (value === null || value === undefined) {
        throw new Error(message || 'Value is null or undefined');
    }
}

try {
    console.log('\n' + '='.repeat(70));
    console.log('SECTION 1: ExcelConnector Basic Operations');
    console.log('='.repeat(70));

    test('Create ExcelConnector instance', () => {
        excel = new ExcelConnector();
        assertNotNull(excel, 'Excel instance should not be null');
    });

    test('Get Excel version', () => {
        const version = excel.getVersion();
        assertNotNull(version, 'Version should not be null');
        console.log(`    Excel version: ${version}`);
    });

    test('Set and get visibility', () => {
        excel.setVisible(true);
        const visible = excel.getVisible();
        assertEquals(visible, true, 'Visibility should be true');
        excel.setVisible(false);
    });

    test('Set display alerts', () => {
        excel.setDisplayAlerts(false);
        console.log('    Display alerts disabled');
    });

    test('Set screen updating', () => {
        excel.setScreenUpdating(false);
        console.log('    Screen updating disabled');
    });

    test('Set calculation mode', () => {
        excel.setCalculation(XlCalculation.xlCalculationManual);
        console.log('    Calculation mode set to manual');
    });

    test('Evaluate formula', () => {
        const result = excel.evaluate('=2+2');
        assertEquals(result, 4, 'Formula evaluation should return 4');
        console.log(`    Evaluated: =2+2 = ${result}`);
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 2: Workbook Operations');
    console.log('='.repeat(70));

    test('Add new workbook', () => {
        workbook = excel.addWorkbook();
        assertNotNull(workbook, 'Workbook should not be null');
    });

    test('Get workbook name', () => {
        const name = workbook.getName();
        assertNotNull(name, 'Workbook name should not be null');
        console.log(`    Workbook name: ${name}`);
    });

    test('Get active workbook', () => {
        const activeWb = excel.getActiveWorkbook();
        assertNotNull(activeWb, 'Active workbook should not be null');
    });

    test('Get worksheet count', () => {
        const count = workbook.getWorksheetCount();
        assert(count > 0, 'Should have at least one worksheet');
        console.log(`    Worksheet count: ${count}`);
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 3: Worksheet Operations');
    console.log('='.repeat(70));

    let sheet1 = null;

    test('Get worksheet by index', () => {
        sheet1 = workbook.getWorksheet(1);
        assertNotNull(sheet1, 'Worksheet should not be null');
    });

    test('Get worksheet name', () => {
        const name = sheet1.getName();
        assertNotNull(name, 'Worksheet name should not be null');
        console.log(`    Worksheet name: ${name}`);
    });

    test('Set worksheet name', () => {
        sheet1.setName('TestSheet');
        const name = sheet1.getName();
        assertEquals(name, 'TestSheet', 'Worksheet name should be TestSheet');
    });

    test('Get worksheet by name', () => {
        const sheet = workbook.getWorksheetByName('TestSheet');
        assertNotNull(sheet, 'Worksheet should not be null');
    });

    test('Add new worksheet', () => {
        const newSheet = workbook.addWorksheet('NewSheet');
        assertNotNull(newSheet, 'New worksheet should not be null');
        const name = newSheet.getName();
        assertEquals(name, 'NewSheet', 'New worksheet name should be NewSheet');
    });

    test('Get active sheet', () => {
        const activeSheet = workbook.getActiveSheet();
        assertNotNull(activeSheet, 'Active sheet should not be null');
    });

    test('Activate worksheet', () => {
        sheet1.activate();
        console.log('    Worksheet activated');
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 4: Range and Cell Operations');
    console.log('='.repeat(70));

    test('Get cell by row/column', () => {
        const cell = sheet1.getCells(1, 1);
        assertNotNull(cell, 'Cell should not be null');
    });

    test('Set and get cell value', () => {
        const cell = sheet1.getCells(1, 1);
        cell.setValue('Hello Excel');
        const value = cell.getValue();
        assertEquals(value, 'Hello Excel', 'Cell value should be "Hello Excel"');
    });

    test('Get range by address', () => {
        const range = sheet1.getRange('A1:C3');
        assertNotNull(range, 'Range should not be null');
    });

    test('Set range values', () => {
        const cell1 = sheet1.getCells(2, 1);
        const cell2 = sheet1.getCells(2, 2);
        const cell3 = sheet1.getCells(2, 3);
        cell1.setValue(10);
        cell2.setValue(20);
        cell3.setValue(30);
        console.log('    Set values: 10, 20, 30');
    });

    test('Set and get formula', () => {
        const cell = sheet1.getCells(3, 1);
        cell.setFormula('=A2+B2+C2');
        const formula = cell.getFormula();
        assertEquals(formula, '=A2+B2+C2', 'Formula should match');
        const value = cell.getValue();
        assertEquals(value, 60, 'Formula result should be 60');
        console.log(`    Formula: ${formula} = ${value}`);
    });

    test('Get cell address', () => {
        const cell = sheet1.getCells(1, 1);
        const address = cell.getAddress();
        assertNotNull(address, 'Address should not be null');
        console.log(`    Cell address: ${address}`);
    });

    test('Get cell row and column', () => {
        const cell = sheet1.getCells(5, 3);
        const row = cell.getRow();
        const col = cell.getColumn();
        assertEquals(row, 5, 'Row should be 5');
        assertEquals(col, 3, 'Column should be 3');
    });

    test('Get used range', () => {
        const usedRange = sheet1.getUsedRange();
        assertNotNull(usedRange, 'Used range should not be null');
        const address = usedRange.getAddress();
        console.log(`    Used range: ${address}`);
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 5: Formatting Operations');
    console.log('='.repeat(70));

    test('Set font properties', () => {
        const cell = sheet1.getCells(1, 1);
        const font = cell.getFont();
        font.setName('Arial');
        font.setSize(14);
        font.setBold(true);
        font.setColor(255); // Red
        console.log('    Font: Arial, 14pt, Bold, Red');
    });

    test('Get font properties', () => {
        const cell = sheet1.getCells(1, 1);
        const font = cell.getFont();
        const name = font.getName();
        const size = font.getSize();
        const bold = font.getBold();
        console.log(`    Font: ${name}, ${size}pt, Bold: ${bold}`);
    });

    test('Set cell background color', () => {
        const cell = sheet1.getCells(1, 1);
        const interior = cell.getInterior();
        interior.setColor(65535); // Yellow
        console.log('    Background color set to yellow');
    });

    test('Set number format', () => {
        const cell = sheet1.getCells(4, 1);
        cell.setValue(1234.56);
        cell.setNumberFormat('$#,##0.00');
        const format = cell.getNumberFormat();
        console.log(`    Number format: ${format}`);
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 6: Range Operations');
    console.log('='.repeat(70));

    test('Copy range', () => {
        const source = sheet1.getRange('A1');
        const dest = sheet1.getRange('D1');
        source.copy(dest);
        const destValue = dest.getValue();
        assertEquals(destValue, 'Hello Excel', 'Copied value should match');
    });

    test('Clear contents', () => {
        const cell = sheet1.getCells(5, 1);
        cell.setValue('Test');
        cell.clearContents();
        const value = cell.getValue();
        assert(value === null || value === '', 'Cell should be empty');
    });

    test('Find value', () => {
        const range = sheet1.getUsedRange();
        const found = range.find('Hello Excel');
        assertNotNull(found, 'Should find the value');
        console.log(`    Found at: ${found.getAddress()}`);
    });

    test('Replace value', () => {
        const range = sheet1.getUsedRange();
        const success = range.replace('Hello Excel', 'Hello World');
        assert(success, 'Replace should succeed');
        const cell = sheet1.getCells(1, 1);
        const value = cell.getValue();
        assertEquals(value, 'Hello World', 'Value should be replaced');
    });

    test('Auto fit columns', () => {
        const range = sheet1.getRange('A:D');
        range.autoFit();
        console.log('    Columns auto-fitted');
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 7: Chart Operations');
    console.log('='.repeat(70));

    let chart = null;

    test('Create chart', () => {
        // Set up data for chart
        const headers = ['Month', 'Sales'];
        sheet1.getCells(6, 1).setValue(headers[0]);
        sheet1.getCells(6, 2).setValue(headers[1]);
        sheet1.getCells(7, 1).setValue('Jan');
        sheet1.getCells(7, 2).setValue(100);
        sheet1.getCells(8, 1).setValue('Feb');
        sheet1.getCells(8, 2).setValue(150);
        sheet1.getCells(9, 1).setValue('Mar');
        sheet1.getCells(9, 2).setValue(200);

        chart = sheet1.addChart();
        assertNotNull(chart, 'Chart should not be null');
    });

    test('Set chart type', () => {
        chart.setChartType(XlChartType.xlColumnClustered);
        const chartType = chart.getChartType();
        assertEquals(chartType, XlChartType.xlColumnClustered, 'Chart type should match');
    });

    test('Set chart data source', () => {
        const dataRange = sheet1.getRange('A6:B9');
        chart.setSourceData(dataRange);
        console.log('    Chart data source set');
    });

    test('Set chart title', () => {
        chart.setTitleText('Monthly Sales');
        assert(chart.getHasTitle(), 'Chart should have title');
        console.log('    Chart title set');
    });

    test('Set chart legend', () => {
        chart.setHasLegend(true);
        assert(chart.getHasLegend(), 'Chart should have legend');
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 8: File Operations');
    console.log('='.repeat(70));

    const testFile = path.join(testDir, 'test-workbook.xlsx');
    const testFileCsv = path.join(testDir, 'test-workbook.csv');
    const testFilePdf = path.join(testDir, 'test-workbook.pdf');

    test('Save workbook', () => {
        workbook.saveAs(testFile, XlFileFormat.xlOpenXMLWorkbook);
        assert(fs.existsSync(testFile), 'File should exist');
        console.log(`    Saved to: ${testFile}`);
    });

    test('Get workbook full name', () => {
        const fullName = workbook.getFullName();
        assertNotNull(fullName, 'Full name should not be null');
        console.log(`    Full name: ${fullName}`);
    });

    test('Get workbook path', () => {
        const wbPath = workbook.getPath();
        assertNotNull(wbPath, 'Path should not be null');
        console.log(`    Path: ${wbPath}`);
    });

    test('Save as CSV', () => {
        workbook.saveAs(testFileCsv, XlFileFormat.xlCSV);
        assert(fs.existsSync(testFileCsv), 'CSV file should exist');
        console.log(`    Saved CSV to: ${testFileCsv}`);
    });

    test('Export as PDF', () => {
        workbook.exportAsPDF(testFilePdf);
        assert(fs.existsSync(testFilePdf), 'PDF file should exist');
        console.log(`    Exported PDF to: ${testFilePdf}`);
    });

    test('Close and reopen workbook', () => {
        workbook.close(false);
        workbook = excel.openWorkbook(testFile, false);
        assertNotNull(workbook, 'Reopened workbook should not be null');
        const sheet = workbook.getWorksheet(1);
        const cell = sheet.getCells(1, 1);
        const value = cell.getValue();
        assertEquals(value, 'Hello World', 'Value should persist after save/load');
    });

    test('Open workbook read-only', () => {
        const wb = excel.openWorkbook(testFile, true);
        assertNotNull(wb, 'Read-only workbook should not be null');
        assert(wb.getReadOnly(), 'Workbook should be read-only');
        wb.close(false);
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 9: Advanced Operations');
    console.log('='.repeat(70));

    test('Calculate workbook', () => {
        excel.setCalculation(XlCalculation.xlCalculationManual);
        const sheet = workbook.getWorksheet(1);
        const cell = sheet.getCells(10, 1);
        cell.setFormula('=RAND()');
        excel.calculate();
        const value = cell.getValue();
        assertNotNull(value, 'Calculated value should not be null');
        console.log(`    Calculated RAND(): ${value}`);
    });

    test('Copy worksheet', () => {
        const sheet = workbook.getWorksheet(1);
        sheet.copy();
        const count = workbook.getWorksheetCount();
        assert(count > 1, 'Should have more than one worksheet after copy');
        console.log(`    Worksheet copied, total sheets: ${count}`);
    });

    test('Delete worksheet', () => {
        const initialCount = workbook.getWorksheetCount();
        if (initialCount > 1) {
            const sheet = workbook.getWorksheet(2);
            sheet.delete();
            const newCount = workbook.getWorksheetCount();
            assertEquals(newCount, initialCount - 1, 'Worksheet count should decrease');
            console.log(`    Worksheet deleted, remaining: ${newCount}`);
        }
    });

    test('Worksheet visibility', () => {
        const sheet = workbook.getWorksheet(1);
        const visible = sheet.getVisible();
        console.log(`    Worksheet visible: ${visible}`);
    });

    test('Get worksheet index', () => {
        const sheet = workbook.getWorksheet(1);
        const index = sheet.getIndex();
        assertEquals(index, 1, 'Index should be 1');
    });

    console.log('\n' + '='.repeat(70));
    console.log('SECTION 10: Cleanup and Resource Management');
    console.log('='.repeat(70));

    test('Set saved flag', () => {
        workbook.setSaved(true);
        const saved = workbook.getSaved();
        assert(saved, 'Workbook should be marked as saved');
    });

    test('Close workbook without saving', () => {
        workbook.close(false);
        console.log('    Workbook closed');
    });

    test('Quit Excel', () => {
        excel.quit();
        console.log('    Excel quit');
    });

    test('Release resources', () => {
        excel.release();
        console.log('    Resources released');
    });

} catch (error) {
    console.error('\n\n✗ FATAL ERROR:', error.message);
    console.error(error.stack);
    testsFailed++;
} finally {
    // Cleanup
    try {
        if (workbook) {
            try { workbook.close(false); } catch (e) {}
        }
        if (excel) {
            try { excel.quit(); } catch (e) {}
            try { excel.release(); } catch (e) {}
        }
    } catch (e) {
        console.error('Error during cleanup:', e.message);
    }
}

console.log('\n' + '='.repeat(70));
console.log('TEST RESULTS');
console.log('='.repeat(70));
console.log(`Total tests: ${testsPassed + testsFailed}`);
console.log(`✓ Passed: ${testsPassed}`);
console.log(`✗ Failed: ${testsFailed}`);
console.log(`Success rate: ${((testsPassed / (testsPassed + testsFailed)) * 100).toFixed(1)}%`);
console.log('='.repeat(70));

if (testsFailed === 0) {
    console.log('\n🎉 ALL TESTS PASSED! 🎉\n');
} else {
    console.log(`\n⚠️  ${testsFailed} TEST(S) FAILED\n`);
}

console.log(`Test files saved to: ${testDir}`);

process.exit(testsFailed > 0 ? 1 : 0);
