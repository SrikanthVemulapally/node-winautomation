/**
 * Comprehensive PowerPoint Connector Test Suite
 * Tests all enhanced methods matching UiPath capabilities
 */

const { 
    PowerPointConnector, 
    PpSaveAsFileType,
    PpSlideLayout,
    MsoAutoShapeType,
    PpParagraphAlignment
} = require('../index.js');
const path = require('path');
const fs = require('fs');

console.log('='.repeat(70));
console.log('COMPREHENSIVE POWERPOINT CONNECTOR TEST SUITE');
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
    const ppt = new PowerPointConnector();
    const testDir = path.join(__dirname, 'powerpoint-test-files');
    
    // Create test directory
    if (!fs.existsSync(testDir)) {
        fs.mkdirSync(testDir, { recursive: true });
    }
    
    const testFile = path.join(testDir, 'test-presentation.pptx');
    const testFilePDF = path.join(testDir, 'test-presentation.pdf');
    
    try {
        ppt.setVisible(false);
        
        console.log('\n--- Testing Application Operations ---');
        
        // Test visibility
        try {
            ppt.setVisible(false);
            const visible = ppt.getVisible();
            logTest('setVisible/getVisible', visible === false || visible === 0);
        } catch (e) {
            logTest('setVisible/getVisible', false, e.message);
        }
        
        // Test version
        try {
            const version = ppt.getVersion();
            logTest('getVersion', typeof version === 'string' && version.length > 0);
        } catch (e) {
            logTest('getVersion', false, e.message);
        }
        
        console.log('\n--- Testing Presentation Operations ---');
        
        let pres;
        
        // Test create presentation
        try {
            pres = ppt.addPresentation();
            logTest('addPresentation', pres !== null);
        } catch (e) {
            logTest('addPresentation', false, e.message);
            ppt.quit();
            ppt.release();
            return;
        }
        
        // Test get active presentation
        try {
            const activePres = ppt.getActivePresentation();
            logTest('getActivePresentation', activePres !== null);
        } catch (e) {
            logTest('getActivePresentation', false, e.message);
        }
        
        // Test presentation count
        try {
            const count = ppt.getPresentationCount();
            logTest('getPresentationCount', count === 1);
        } catch (e) {
            logTest('getPresentationCount', false, e.message);
        }
        
        // Test presentation properties
        try {
            const name = pres.getName();
            logTest('Presentation getName', name.length > 0);
        } catch (e) {
            logTest('Presentation getName', false, e.message);
        }
        
        console.log('\n--- Testing Slide Operations ---');
        
        // Test initial slide count
        try {
            const count = pres.getSlideCount();
            logTest('getSlideCount (initial)', count >= 0);
        } catch (e) {
            logTest('getSlideCount (initial)', false, e.message);
        }
        
        // Test add slide
        try {
            const slide = pres.addSlide(1, PpSlideLayout.ppLayoutBlank);
            logTest('addSlide', slide !== null);
        } catch (e) {
            logTest('addSlide', false, e.message);
        }
        
        // Test get slide
        try {
            const slide = pres.getSlide(1);
            logTest('getSlide', slide !== null);
        } catch (e) {
            logTest('getSlide', false, e.message);
        }
        
        // Test slide index
        try {
            const slide = pres.getSlide(1);
            const index = slide.getSlideIndex();
            logTest('getSlideIndex', index === 1);
        } catch (e) {
            logTest('getSlideIndex', false, e.message);
        }
        
        console.log('\n--- Testing Shape Operations ---');
        
        let slide;
        try {
            slide = pres.getSlide(1);
        } catch (e) {
            console.log('Cannot get slide for shape tests');
        }
        
        if (slide) {
            // Test add text box
            try {
                const textBox = slide.addTextBox(100, 100, 300, 100);
                logTest('addTextBox', textBox !== null);
            } catch (e) {
                logTest('addTextBox', false, e.message);
            }
            
            // Test shape count
            try {
                const count = slide.getShapeCount();
                logTest('getShapeCount', count >= 1);
            } catch (e) {
                logTest('getShapeCount', false, e.message);
            }
            
            // Test get shape
            try {
                const shape = slide.getShape(1);
                logTest('getShape', shape !== null);
            } catch (e) {
                logTest('getShape', false, e.message);
            }
            
            // Test add shape
            try {
                const shape = slide.addShape(MsoAutoShapeType.msoShapeRectangle, 50, 50, 200, 100);
                logTest('addShape', shape !== null);
            } catch (e) {
                logTest('addShape', false, e.message);
            }
        }
        
        console.log('\n--- Testing Text Operations ---');
        
        if (slide) {
            // Test set/get text
            try {
                const textBox = slide.addTextBox(100, 200, 300, 100);
                textBox.setText('Hello PowerPoint');
                const text = textBox.getText();
                logTest('setText/getText', text.includes('Hello PowerPoint'));
            } catch (e) {
                logTest('setText/getText', false, e.message);
            }
            
            // Test text formatting
            try {
                const textBox = slide.addTextBox(100, 320, 300, 100);
                textBox.setText('Formatted Text');
                const textFrame = textBox.getTextFrame();
                const textRange = textFrame.getTextRange();
                const font = textRange.getFont();
                font.setName('Arial');
                font.setSize(24);
                font.setBold(true);
                font.setItalic(true);
                logTest('Text formatting', true);
            } catch (e) {
                logTest('Text formatting', false, e.message);
            }
            
            // Test paragraph alignment
            try {
                const textBox = slide.addTextBox(100, 440, 300, 100);
                textBox.setText('Centered Text');
                const textFrame = textBox.getTextFrame();
                const textRange = textFrame.getTextRange();
                const paraFormat = textRange.getParagraphFormat();
                paraFormat.setAlignment(PpParagraphAlignment.ppAlignCenter);
                logTest('Paragraph alignment', true);
            } catch (e) {
                logTest('Paragraph alignment', false, e.message);
            }
        }
        
        console.log('\n--- Testing Shape Properties ---');
        
        if (slide) {
            // Test shape position and size
            try {
                const shape = slide.addShape(MsoAutoShapeType.msoShapeOval, 400, 100, 150, 150);
                shape.setLeft(420);
                shape.setTop(120);
                shape.setWidth(180);
                shape.setHeight(180);
                const left = shape.getLeft();
                const top = shape.getTop();
                logTest('Shape position/size', left === 420 && top === 120);
            } catch (e) {
                logTest('Shape position/size', false, e.message);
            }
            
            // Test shape name
            try {
                const shape = slide.getShape(1);
                shape.setName('TestShape');
                const name = shape.getName();
                logTest('Shape name', name === 'TestShape');
            } catch (e) {
                logTest('Shape name', false, e.message);
            }
        }
        
        console.log('\n--- Testing Table Operations ---');
        
        if (slide) {
            // Test add table
            try {
                const tableShape = slide.addTable(3, 3, 50, 400, 400, 150);
                logTest('addTable', tableShape !== null);
            } catch (e) {
                logTest('addTable', false, e.message);
            }
            
            // Test table cell operations
            try {
                const shapes = slide.getShapes();
                let tableShape = null;
                const count = slide.getShapeCount();
                for (let i = 1; i <= count; i++) {
                    const shape = slide.getShape(i);
                    if (shape.hasTable()) {
                        tableShape = shape;
                        break;
                    }
                }
                
                if (tableShape) {
                    const table = tableShape.getTable();
                    table.setCellText(1, 1, 'Cell 1,1');
                    table.setCellText(1, 2, 'Cell 1,2');
                    const text = table.getCellText(1, 1);
                    logTest('Table cell operations', text.includes('Cell 1,1'));
                } else {
                    logTest('Table cell operations', false, 'No table found');
                }
            } catch (e) {
                logTest('Table cell operations', false, e.message);
            }
            
            // Test table row/column count
            try {
                const shapes = slide.getShapes();
                let tableShape = null;
                const count = slide.getShapeCount();
                for (let i = 1; i <= count; i++) {
                    const shape = slide.getShape(i);
                    if (shape.hasTable()) {
                        tableShape = shape;
                        break;
                    }
                }
                
                if (tableShape) {
                    const table = tableShape.getTable();
                    const rowCount = table.getRowCount();
                    const colCount = table.getColumnCount();
                    logTest('Table row/column count', rowCount === 3 && colCount === 3);
                } else {
                    logTest('Table row/column count', false, 'No table found');
                }
            } catch (e) {
                logTest('Table row/column count', false, e.message);
            }
        }
        
        console.log('\n--- Testing Slide Notes ---');
        
        if (slide) {
            // Test set/get notes
            try {
                slide.setNotes('These are speaker notes');
                const notes = slide.getNotes();
                logTest('Slide notes', notes.includes('speaker notes'));
            } catch (e) {
                logTest('Slide notes', false, e.message);
            }
        }
        
        console.log('\n--- Testing Multiple Slides ---');
        
        // Test add multiple slides
        try {
            pres.addSlide(2, PpSlideLayout.ppLayoutTitle);
            pres.addSlide(3, PpSlideLayout.ppLayoutText);
            const count = pres.getSlideCount();
            logTest('Add multiple slides', count >= 3);
        } catch (e) {
            logTest('Add multiple slides', false, e.message);
        }
        
        console.log('\n--- Testing Save Operations ---');
        
        // Test save as
        try {
            pres.saveAs(testFile, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            logTest('saveAs', fs.existsSync(testFile));
        } catch (e) {
            logTest('saveAs', false, e.message);
        }
        
        // Test save status
        try {
            const saved = pres.getSaved();
            logTest('getSaved', saved === true || saved === -1);
        } catch (e) {
            logTest('getSaved', false, e.message);
        }
        
        // Test export to PDF
        try {
            pres.exportToPDF(testFilePDF);
            logTest('exportToPDF', fs.existsSync(testFilePDF));
        } catch (e) {
            logTest('exportToPDF', false, e.message);
        }
        
        // Test full name
        try {
            const fullName = pres.getFullName();
            logTest('getFullName', fullName.length > 0);
        } catch (e) {
            logTest('getFullName', false, e.message);
        }
        
        console.log('\n--- Testing Close and Open ---');
        
        // Test close presentation
        try {
            pres.close();
            logTest('close presentation', true);
        } catch (e) {
            logTest('close presentation', false, e.message);
        }
        
        // Test open presentation
        try {
            const openedPres = ppt.openPresentation(testFile);
            logTest('openPresentation', openedPres !== null);
            openedPres.close();
        } catch (e) {
            logTest('openPresentation', false, e.message);
        }
        
    } catch (error) {
        console.error('\n❌ Test suite error:', error.message);
        console.error(error.stack);
    } finally {
        // Cleanup
        try {
            ppt.quit();
            ppt.release();
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
