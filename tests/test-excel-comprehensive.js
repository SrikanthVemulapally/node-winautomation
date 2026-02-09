/**
 * Comprehensive Excel Connector Test Suite
 * Tests all enhanced methods matching UiPath capabilities
 */

const { 
    ExcelConnector, 
    XlFileFormat, 
    XlHAlign, 
    XlVAlign, 
    XlBordersIndex, 
    XlLineStyle, 
    XlBorderWeight,
    XlSortOrder 
} = require('../index.js');
const path = require('path');
const fs = require('fs');

console.log('='.repeat(70));
console.log('COMPREHENSIVE EXCEL CONNECTOR TEST SUITE');
console.log('='.repeat(70));

let testsPassed = 0;
let testsFailed = 0;
const testResults = [];

function logTest(testName, passed, error = null) {
    if (passed) {
        console.log(`✓ ${testName}`);
        testsPassed++;
        testResults.push({ test: testName, status: 'PASS' });
    } else {
        console.log(`✗ ${testName}: ${error}`);
        testsFailed++;
        testResults.push({ test: testName, status: 'FAIL', error: error });
    }
}

async function testCellOperations() {
    console.log('\n--- Testing Cell Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        
        // Test writeCell
        try {
            sheet.writeCell(1, 1, 'Test Value');
            const value = sheet.readCell(1, 1);
            logTest('writeCell/readCell', value === 'Test Value');
        } catch (e) {
            logTest('writeCell/readCell', false, e.message);
        }
        
        // Test writeCell with number
        try {
            sheet.writeCell(2, 1, 42);
            const value = sheet.readCell(2, 1);
            logTest('writeCell (number)', value === 42);
        } catch (e) {
            logTest('writeCell (number)', false, e.message);
        }
        
        // Test writeCell with formula
        try {
            sheet.writeCell(3, 1, 10);
            sheet.writeCell(3, 2, 20);
            sheet.writeCell(3, 3, '=A3+B3');
            const value = sheet.readCell(3, 3);
            logTest('writeCell (formula)', value === 30);
        } catch (e) {
            logTest('writeCell (formula)', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function testRangeOperations() {
    console.log('\n--- Testing Range Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        
        // Test writeRange/readRange
        try {
            const data = [
                ['Name', 'Age', 'City'],
                ['John', 30, 'NYC'],
                ['Jane', 25, 'LA']
            ];
            sheet.writeRange('A1:C3', data);
            const readData = sheet.readRange('A1:C3');
            logTest('writeRange/readRange', readData && readData.length === 3);
        } catch (e) {
            logTest('writeRange/readRange', false, e.message);
        }
        
        // Test appendRow
        try {
            sheet.writeCell(1, 1, 'Item');
            sheet.writeCell(2, 1, 'Apple');
            sheet.appendRow(['Banana']);
            const value = sheet.readCell(3, 1);
            logTest('appendRow', value === 'Banana');
        } catch (e) {
            logTest('appendRow', false, e.message);
        }
        
        // Test getLastRow
        try {
            const lastRow = sheet.getLastRow();
            logTest('getLastRow', lastRow >= 3);
        } catch (e) {
            logTest('getLastRow', false, e.message);
        }
        
        // Test getLastColumn
        try {
            const lastCol = sheet.getLastColumn();
            logTest('getLastColumn', lastCol >= 1);
        } catch (e) {
            logTest('getLastColumn', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function testFormattingOperations() {
    console.log('\n--- Testing Formatting Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        
        // Test font formatting
        try {
            sheet.writeCell(1, 1, 'Bold Text');
            const range = sheet.getCells(1, 1);
            const font = range.getFont();
            font.setBold(true);
            font.setSize(14);
            font.setName('Arial');
            logTest('Font formatting', true);
        } catch (e) {
            logTest('Font formatting', false, e.message);
        }
        
        // Test cell color
        try {
            sheet.writeCell(2, 1, 'Colored Cell');
            const range = sheet.getCells(2, 1);
            const interior = range.getInterior();
            interior.setColor(0xFFFF00); // Yellow
            logTest('Cell color', true);
        } catch (e) {
            logTest('Cell color', false, e.message);
        }
        
        // Test borders
        try {
            sheet.writeCell(3, 1, 'Bordered Cell');
            const range = sheet.getCells(3, 1);
            range.setBorder(XlBordersIndex.xlEdgeLeft, XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            range.setBorder(XlBordersIndex.xlEdgeTop, XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            logTest('Borders', true);
        } catch (e) {
            logTest('Borders', false, e.message);
        }
        
        // Test alignment
        try {
            sheet.writeCell(4, 1, 'Centered');
            const range = sheet.getCells(4, 1);
            range.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
            range.setVerticalAlignment(XlVAlign.xlVAlignCenter);
            logTest('Alignment', true);
        } catch (e) {
            logTest('Alignment', false, e.message);
        }
        
        // Test wrap text
        try {
            sheet.writeCell(5, 1, 'This is a long text that should wrap');
            const range = sheet.getCells(5, 1);
            range.setWrapText(true);
            logTest('Wrap text', true);
        } catch (e) {
            logTest('Wrap text', false, e.message);
        }
        
        // Test number format
        try {
            sheet.writeCell(6, 1, 1234.56);
            const range = sheet.getCells(6, 1);
            range.setNumberFormat('$#,##0.00');
            logTest('Number format', true);
        } catch (e) {
            logTest('Number format', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function testColumnRowOperations() {
    console.log('\n--- Testing Column/Row Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        
        // Test setColumnWidth
        try {
            sheet.setColumnWidth(1, 20);
            logTest('setColumnWidth', true);
        } catch (e) {
            logTest('setColumnWidth', false, e.message);
        }
        
        // Test setRowHeight
        try {
            sheet.setRowHeight(1, 30);
            logTest('setRowHeight', true);
        } catch (e) {
            logTest('setRowHeight', false, e.message);
        }
        
        // Test autoFitColumns
        try {
            sheet.writeCell(1, 1, 'This is a long text');
            sheet.autoFitColumns('A:A');
            logTest('autoFitColumns', true);
        } catch (e) {
            logTest('autoFitColumns', false, e.message);
        }
        
        // Test autoFitRows
        try {
            sheet.autoFitRows('1:1');
            logTest('autoFitRows', true);
        } catch (e) {
            logTest('autoFitRows', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function testWorksheetOperations() {
    console.log('\n--- Testing Worksheet Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        
        // Test addWorksheet
        try {
            const sheet2 = workbook.addWorksheet('TestSheet');
            sheet2.setName('RenamedSheet');
            logTest('addWorksheet and rename', sheet2.getName() === 'RenamedSheet');
        } catch (e) {
            logTest('addWorksheet and rename', false, e.message);
        }
        
        // Test getWorksheetCount
        try {
            const count = workbook.getWorksheetCount();
            logTest('getWorksheetCount', count >= 2);
        } catch (e) {
            logTest('getWorksheetCount', false, e.message);
        }
        
        // Test getWorksheet by index
        try {
            const sheet = workbook.getWorksheet(1);
            logTest('getWorksheet by index', sheet !== null);
        } catch (e) {
            logTest('getWorksheet by index', false, e.message);
        }
        
        // Test getWorksheetByName
        try {
            const sheet = workbook.getWorksheetByName('RenamedSheet');
            logTest('getWorksheetByName', sheet.getName() === 'RenamedSheet');
        } catch (e) {
            logTest('getWorksheetByName', false, e.message);
        }
        
        // Test worksheet visibility
        try {
            const sheet = workbook.getWorksheet(2);
            sheet.setVisible(false);
            logTest('setVisible (worksheet)', true);
        } catch (e) {
            logTest('setVisible (worksheet)', false, e.message);
        }
        
        // Test protect/unprotect
        try {
            const sheet = workbook.getActiveSheet();
            sheet.protect('test123');
            sheet.unprotect('test123');
            logTest('protect/unprotect worksheet', true);
        } catch (e) {
            logTest('protect/unprotect worksheet', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function testRangeAdvanced() {
    console.log('\n--- Testing Advanced Range Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        
        // Test merge/unmerge
        try {
            sheet.writeCell(1, 1, 'Merged Cell');
            const range = sheet.getRange('A1:C1');
            range.merge();
            range.unmerge();
            logTest('merge/unmerge cells', true);
        } catch (e) {
            logTest('merge/unmerge cells', false, e.message);
        }
        
        // Test find
        try {
            sheet.writeCell(2, 1, 'FindMe');
            const range = sheet.getUsedRange();
            const found = range.find('FindMe');
            logTest('find value', found !== null);
        } catch (e) {
            logTest('find value', false, e.message);
        }
        
        // Test replace
        try {
            sheet.writeCell(3, 1, 'OldValue');
            const range = sheet.getUsedRange();
            range.replace('OldValue', 'NewValue');
            const value = sheet.readCell(3, 1);
            logTest('replace value', value === 'NewValue');
        } catch (e) {
            logTest('replace value', false, e.message);
        }
        
        // Test clear operations
        try {
            sheet.writeCell(4, 1, 'ClearMe');
            const range = sheet.getCells(4, 1);
            range.clearContents();
            const value = sheet.readCell(4, 1);
            logTest('clearContents', value === null || value === '');
        } catch (e) {
            logTest('clearContents', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function testSortOperations() {
    console.log('\n--- Testing Sort Operations ---');
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        excel.setDisplayAlerts(false);
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        
        // Test sortRange
        try {
            sheet.writeCell(1, 1, 'Name');
            sheet.writeCell(2, 1, 'Charlie');
            sheet.writeCell(3, 1, 'Alice');
            sheet.writeCell(4, 1, 'Bob');
            
            sheet.sortRange('A1:A4', 'A1', 1); // Ascending
            logTest('sortRange', true);
        } catch (e) {
            logTest('sortRange', false, e.message);
        }
        
        workbook.close(false);
    } finally {
        excel.quit();
        excel.release();
    }
}

async function runAllTests() {
    try {
        await testCellOperations();
        await testRangeOperations();
        await testFormattingOperations();
        await testColumnRowOperations();
        await testWorksheetOperations();
        await testRangeAdvanced();
        await testSortOperations();
        
        console.log('\n' + '='.repeat(70));
        console.log('TEST SUMMARY');
        console.log('='.repeat(70));
        console.log(`Total Tests: ${testsPassed + testsFailed}`);
        console.log(`✓ Passed: ${testsPassed}`);
        console.log(`✗ Failed: ${testsFailed}`);
        console.log(`Success Rate: ${((testsPassed / (testsPassed + testsFailed)) * 100).toFixed(1)}%`);
        
        if (testsFailed > 0) {
            console.log('\nFailed Tests:');
            testResults.filter(r => r.status === 'FAIL').forEach(r => {
                console.log(`  - ${r.test}: ${r.error}`);
            });
        }
        
        console.log('='.repeat(70));
        
    } catch (error) {
        console.error('\n❌ Test suite error:', error.message);
        console.error(error.stack);
    }
}

runAllTests();
