/**
 * Excel Automation Example
 * 
 * Demonstrates basic Excel automation using the Excel connector.
 * Only includes methods that have been verified to work.
 * 
 * Prerequisites:
 * - Microsoft Excel installed
 * - Build the addon first: npm run build
 */

// Standard professional import style
const { ExcelConnector, XlFileFormat } = require('../../../index.js');
const path = require('path');

console.log('Excel Connector Examples\n');
console.log('='.repeat(50) + '\n');

// Example 1: Create Simple Workbook
async function example1() {
    console.log('=== Example 1: Create Simple Workbook ===\n');
    
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(true);
        
        const workbook = excel.addWorkbook();
        console.log('✓ Workbook created');
        
        const sheet = workbook.getActiveSheet();
        sheet.setName('Sales Data');
        console.log('✓ Sheet renamed to "Sales Data"');
        
        const outputPath = path.join(process.cwd(), 'example1-workbook.xlsx');
        workbook.saveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Workbook saved to: ${outputPath}`);
        
        workbook.close(false);
        console.log('✓ Workbook closed\n');
        
    } finally {
        excel.quit();
        excel.release();
    }
}

// Example 2: Multiple Sheets
async function example2() {
    console.log('=== Example 2: Multiple Sheets ===\n');
    
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(true);
        
        const workbook = excel.addWorkbook();
        
        const sheet1 = workbook.getActiveSheet();
        sheet1.setName('Q1 Data');
        console.log('✓ First sheet: Q1 Data');
        
        const sheet2 = workbook.addWorksheet();
        sheet2.setName('Q2 Data');
        console.log('✓ Second sheet: Q2 Data');
        
        const sheet3 = workbook.addWorksheet();
        sheet3.setName('Q3 Data');
        console.log('✓ Third sheet: Q3 Data');
        
        const outputPath = path.join(process.cwd(), 'example2-multiple-sheets.xlsx');
        workbook.saveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Workbook saved to: ${outputPath}`);
        
        workbook.close(false);
        console.log('✓ Workbook closed\n');
        
    } finally {
        excel.quit();
        excel.release();
    }
}

// Example 3: Display Settings
async function example3() {
    console.log('=== Example 3: Display Settings ===\n');
    
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(false);
        console.log('✓ Excel started (hidden)');
        
        excel.setDisplayAlerts(false);
        console.log('✓ Display alerts disabled');
        
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        sheet.setName('Data');
        console.log('✓ Workbook created');
        
        excel.setVisible(true);
        console.log('✓ Excel now visible');
        
        const outputPath = path.join(process.cwd(), 'example3-display-settings.xlsx');
        workbook.saveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Workbook saved to: ${outputPath}`);
        
        workbook.close(false);
        console.log('✓ Workbook closed\n');
        
    } finally {
        excel.quit();
        excel.release();
    }
}

// Example 4: Save in Different Formats
async function example4() {
    console.log('=== Example 4: Save in Different Formats ===\n');
    
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(true);
        const workbook = excel.addWorkbook();
        const sheet = workbook.getActiveSheet();
        sheet.setName('Export Data');
        
        const xlsxPath = path.join(process.cwd(), 'example4-data.xlsx');
        workbook.saveAs(xlsxPath, XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Saved as XLSX: ${xlsxPath}`);
        
        const csvPath = path.join(process.cwd(), 'example4-data.csv');
        workbook.saveAs(csvPath, XlFileFormat.xlCSV);
        console.log(`✓ Saved as CSV: ${csvPath}`);
        
        const htmlPath = path.join(process.cwd(), 'example4-data.html');
        workbook.saveAs(htmlPath, XlFileFormat.xlHtml);
        console.log(`✓ Saved as HTML: ${htmlPath}`);
        
        workbook.close(false);
        console.log('✓ Workbook closed\n');
        
    } finally {
        excel.quit();
        excel.release();
    }
}

// Example 5: Get Excel Version
async function example5() {
    console.log('=== Example 5: Get Excel Version ===\n');
    
    const excel = new ExcelConnector();
    
    try {
        const version = excel.getVersion();
        console.log(`Excel Version: ${version}`);
        console.log('(16.0 = Excel 2016/2019/365, 15.0 = Excel 2013, 14.0 = Excel 2010)\n');
        
    } finally {
        excel.release();
    }
}

// Run all examples
async function runAll() {
    try {
        await example1();
        await example2();
        await example3();
        await example4();
        await example5();
        
        console.log('='.repeat(50));
        console.log('✅ All examples completed successfully!\n');
        console.log('Generated files:');
        console.log('  - example1-workbook.xlsx');
        console.log('  - example2-multiple-sheets.xlsx');
        console.log('  - example3-display-settings.xlsx');
        console.log('  - example4-data.xlsx, .csv, .html');
        
    } catch (error) {
        console.error('\n❌ Error:', error.message);
        console.error(error.stack);
    }
}

runAll();
