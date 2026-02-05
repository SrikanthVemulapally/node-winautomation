/**
 * Excel Automation Example - Standard Import Style
 * 
 * Demonstrates professional, standard import syntax
 */

// STANDARD PROFESSIONAL STYLE - Direct class imports
const { ExcelConnector, XlFileFormat } = require('../../../index.js');
const path = require('path');

async function example() {
    console.log('Excel Example - Standard Import Style\n');
    
    // Clean, professional instantiation
    const excel = new ExcelConnector();
    
    try {
        excel.setVisible(true);
        
        const workbook = excel.addWorkbook();
        console.log('✓ Workbook created');
        
        const sheet = workbook.getActiveSheet();
        sheet.setName('Sales Data');
        console.log('✓ Sheet renamed');
        
        const outputPath = path.join(process.cwd(), 'standard-style-workbook.xlsx');
        workbook.saveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
        console.log(`✓ Saved: ${outputPath}`);
        
        workbook.close(false);
        
    } finally {
        excel.quit();
        excel.release();
    }
}

example().catch(console.error);
