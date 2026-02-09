/**
 * Comprehensive Word Connector Test Suite
 * Tests all enhanced methods matching UiPath capabilities
 */

const { 
    WordConnector, 
    WdSaveFormat, 
    WdParagraphAlignment,
    WdOrientation,
    WdUnderline,
    WdColor 
} = require('../index.js');
const path = require('path');
const fs = require('fs');

console.log('='.repeat(70));
console.log('COMPREHENSIVE WORD CONNECTOR TEST SUITE');
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

async function runTests() {
    const word = new WordConnector();
    const testDir = path.join(__dirname, 'word-test-files');
    
    // Create test directory
    if (!fs.existsSync(testDir)) {
        fs.mkdirSync(testDir, { recursive: true });
    }
    
    const testFile = path.join(testDir, 'test-document.docx');
    const testFilePDF = path.join(testDir, 'test-document.pdf');
    
    try {
        word.setVisible(false);
        word.setDisplayAlerts(false);
        
        console.log('\n--- Testing Application Operations ---');
        
        // Test visibility
        try {
            word.setVisible(false);
            const visible = word.getVisible();
            logTest('setVisible/getVisible', visible === false);
        } catch (e) {
            logTest('setVisible/getVisible', false, e.message);
        }
        
        // Test version
        try {
            const version = word.getVersion();
            logTest('getVersion', typeof version === 'string' && version.length > 0);
        } catch (e) {
            logTest('getVersion', false, e.message);
        }
        
        console.log('\n--- Testing Document Operations ---');
        
        let doc;
        
        // Test create document
        try {
            doc = word.addDocument();
            logTest('addDocument', doc !== null);
        } catch (e) {
            logTest('addDocument', false, e.message);
            word.quit();
            word.release();
            return;
        }
        
        // Test get active document
        try {
            const activeDoc = word.getActiveDocument();
            logTest('getActiveDocument', activeDoc !== null);
        } catch (e) {
            logTest('getActiveDocument', false, e.message);
        }
        
        // Test document count
        try {
            const count = word.getDocumentCount();
            logTest('getDocumentCount', count === 1);
        } catch (e) {
            logTest('getDocumentCount', false, e.message);
        }
        
        console.log('\n--- Testing Text Operations ---');
        
        // Test set/get text
        try {
            doc.setText('Hello World');
            const text = doc.getText();
            logTest('setText/getText', text.includes('Hello World'));
        } catch (e) {
            logTest('setText/getText', false, e.message);
        }
        
        // Test content range
        try {
            const content = doc.getContent();
            content.setText('Test Content');
            const text = content.getText();
            logTest('getContent and Range setText/getText', text.includes('Test Content'));
        } catch (e) {
            logTest('getContent and Range setText/getText', false, e.message);
        }
        
        // Test insert after
        try {
            const content = doc.getContent();
            content.insertAfter(' - Appended');
            const text = doc.getText();
            logTest('insertAfter', text.includes('Appended'));
        } catch (e) {
            logTest('insertAfter', false, e.message);
        }
        
        console.log('\n--- Testing Paragraph Operations ---');
        
        // Test add paragraph
        try {
            doc.setText('First paragraph');
            doc.addParagraph('Second paragraph');
            const count = doc.getParagraphCount();
            logTest('addParagraph', count >= 2);
        } catch (e) {
            logTest('addParagraph', false, e.message);
        }
        
        // Test get paragraph
        try {
            const para = doc.getParagraph(1);
            logTest('getParagraph', para !== null);
        } catch (e) {
            logTest('getParagraph', false, e.message);
        }
        
        // Test paragraph alignment
        try {
            const para = doc.getParagraph(1);
            para.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
            const alignment = para.getAlignment();
            logTest('Paragraph alignment', alignment === WdParagraphAlignment.wdAlignParagraphCenter);
        } catch (e) {
            logTest('Paragraph alignment', false, e.message);
        }
        
        console.log('\n--- Testing Formatting Operations ---');
        
        // Test font formatting
        try {
            doc.setText('Formatted text');
            const content = doc.getContent();
            const font = content.getFont();
            font.setName('Arial');
            font.setSize(14);
            font.setBold(true);
            font.setItalic(true);
            logTest('Font formatting', true);
        } catch (e) {
            logTest('Font formatting', false, e.message);
        }
        
        // Test font color
        try {
            const content = doc.getContent();
            const font = content.getFont();
            font.setColor(WdColor.wdColorRed);
            logTest('Font color', true);
        } catch (e) {
            logTest('Font color', false, e.message);
        }
        
        // Test underline
        try {
            const content = doc.getContent();
            content.setUnderline(WdUnderline.wdUnderlineSingle);
            logTest('Underline', true);
        } catch (e) {
            logTest('Underline', false, e.message);
        }
        
        // Test bold/italic on range
        try {
            const content = doc.getContent();
            content.setBold(true);
            content.setItalic(true);
            logTest('Bold/Italic on range', true);
        } catch (e) {
            logTest('Bold/Italic on range', false, e.message);
        }
        
        console.log('\n--- Testing Table Operations ---');
        
        // Test add table
        try {
            doc.setText('Before table');
            const table = doc.addTable(3, 3);
            logTest('addTable', table !== null);
        } catch (e) {
            logTest('addTable', false, e.message);
        }
        
        // Test table count
        try {
            const count = doc.getTableCount();
            logTest('getTableCount', count === 1);
        } catch (e) {
            logTest('getTableCount', false, e.message);
        }
        
        // Test get table
        try {
            const table = doc.getTable(1);
            logTest('getTable', table !== null);
        } catch (e) {
            logTest('getTable', false, e.message);
        }
        
        // Test table row/column count
        try {
            const table = doc.getTable(1);
            const rowCount = table.getRowCount();
            const colCount = table.getColumnCount();
            logTest('Table row/column count', rowCount === 3 && colCount === 3);
        } catch (e) {
            logTest('Table row/column count', false, e.message);
        }
        
        // Test set/get cell text
        try {
            const table = doc.getTable(1);
            table.setCellText(1, 1, 'Cell 1,1');
            table.setCellText(1, 2, 'Cell 1,2');
            const text = table.getCellText(1, 1);
            logTest('setCellText/getCellText', text.includes('Cell 1,1'));
        } catch (e) {
            logTest('setCellText/getCellText', false, e.message);
        }
        
        // Test add row/column
        try {
            const table = doc.getTable(1);
            table.addRow();
            table.addColumn();
            const rowCount = table.getRowCount();
            const colCount = table.getColumnCount();
            logTest('addRow/addColumn', rowCount === 4 && colCount === 4);
        } catch (e) {
            logTest('addRow/addColumn', false, e.message);
        }
        
        console.log('\n--- Testing Bookmark Operations ---');
        
        // Test add bookmark
        try {
            doc.setText('Bookmarked text');
            const content = doc.getContent();
            doc.addBookmark('TestBookmark', content);
            logTest('addBookmark', true);
        } catch (e) {
            logTest('addBookmark', false, e.message);
        }
        
        // Test bookmark exists
        try {
            const exists = doc.bookmarkExists('TestBookmark');
            logTest('bookmarkExists', exists === true);
        } catch (e) {
            logTest('bookmarkExists', false, e.message);
        }
        
        // Test go to bookmark
        try {
            const range = doc.goToBookmark('TestBookmark');
            logTest('goToBookmark', range !== null);
        } catch (e) {
            logTest('goToBookmark', false, e.message);
        }
        
        console.log('\n--- Testing Page Setup Operations ---');
        
        // Test page setup
        try {
            const pageSetup = doc.getPageSetup();
            pageSetup.setTopMargin(72);
            pageSetup.setBottomMargin(72);
            pageSetup.setLeftMargin(72);
            pageSetup.setRightMargin(72);
            logTest('Page margins', true);
        } catch (e) {
            logTest('Page margins', false, e.message);
        }
        
        // Test orientation
        try {
            const pageSetup = doc.getPageSetup();
            pageSetup.setOrientation(WdOrientation.wdOrientLandscape);
            const orientation = pageSetup.getOrientation();
            logTest('Page orientation', orientation === WdOrientation.wdOrientLandscape);
        } catch (e) {
            logTest('Page orientation', false, e.message);
        }
        
        console.log('\n--- Testing Find/Replace Operations ---');
        
        // Test find
        try {
            doc.setText('Find this text');
            const content = doc.getContent();
            const found = content.find('Find');
            logTest('Find text', found === true);
        } catch (e) {
            logTest('Find text', false, e.message);
        }
        
        // Test replace
        try {
            doc.setText('Replace this');
            const content = doc.getContent();
            content.replace('this', 'that');
            const text = doc.getText();
            logTest('Replace text', text.includes('that'));
        } catch (e) {
            logTest('Replace text', false, e.message);
        }
        
        console.log('\n--- Testing Save Operations ---');
        
        // Test save
        try {
            doc.saveAs(testFile, WdSaveFormat.wdFormatDocx);
            logTest('saveAs', fs.existsSync(testFile));
        } catch (e) {
            logTest('saveAs', false, e.message);
        }
        
        // Test save status
        try {
            const saved = doc.getSaved();
            logTest('getSaved', saved === true);
        } catch (e) {
            logTest('getSaved', false, e.message);
        }
        
        // Test export to PDF
        try {
            doc.exportToPDF(testFilePDF);
            logTest('exportToPDF', fs.existsSync(testFilePDF));
        } catch (e) {
            logTest('exportToPDF', false, e.message);
        }
        
        // Test document properties
        try {
            const name = doc.getName();
            const fullName = doc.getFullName();
            logTest('Document name properties', name.length > 0 && fullName.length > 0);
        } catch (e) {
            logTest('Document name properties', false, e.message);
        }
        
        console.log('\n--- Testing Comment Operations ---');
        
        // Test add comment
        try {
            doc.setText('Commented text');
            const content = doc.getContent();
            doc.addComment(content, 'This is a test comment');
            const count = doc.getCommentCount();
            logTest('addComment', count === 1);
        } catch (e) {
            logTest('addComment', false, e.message);
        }
        
        console.log('\n--- Testing Advanced Operations ---');
        
        // Test section count
        try {
            const count = doc.getSectionCount();
            logTest('getSectionCount', count >= 1);
        } catch (e) {
            logTest('getSectionCount', false, e.message);
        }
        
        // Test update fields
        try {
            doc.updateFields();
            logTest('updateFields', true);
        } catch (e) {
            logTest('updateFields', false, e.message);
        }
        
        // Test close document
        try {
            doc.close(false);
            logTest('close document', true);
        } catch (e) {
            logTest('close document', false, e.message);
        }
        
        console.log('\n--- Testing Open Document ---');
        
        // Test open document
        try {
            const openedDoc = word.openDocument(testFile);
            logTest('openDocument', openedDoc !== null);
            openedDoc.close(false);
        } catch (e) {
            logTest('openDocument', false, e.message);
        }
        
    } catch (error) {
        console.error('\n❌ Test suite error:', error.message);
        console.error(error.stack);
    } finally {
        // Cleanup
        try {
            word.quit(false);
            word.release();
        } catch (e) {
            console.error('Cleanup error:', e.message);
        }
        
        // Print summary
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
    }
}

runTests();
