/**
 * @fileoverview Microsoft PowerPoint Connector
 * Provides a high-level, type-safe API for PowerPoint automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * PowerPoint save formats
 */
const PpSaveAsFileType = {
    ppSaveAsPresentation: 1,
    ppSaveAsTemplate: 5,
    ppSaveAsPDF: 32,
    ppSaveAsOpenXMLPresentation: 24,
    ppSaveAsOpenXMLPresentationMacroEnabled: 25,
    ppSaveAsShow: 7,
    ppSaveAsOpenXMLShow: 28,
};

/**
 * Slide layout constants
 */
const PpSlideLayout = {
    ppLayoutBlank: 12,
    ppLayoutTitle: 1,
    ppLayoutText: 2,
    ppLayoutTwoColumnText: 3,
    ppLayoutTable: 4,
    ppLayoutChart: 8,
    ppLayoutOrgchart: 7,
    ppLayoutObject: 16,
};

/**
 * Shape type constants
 */
const MsoShapeType = {
    msoAutoShape: 1,
    msoTextBox: 17,
    msoPicture: 13,
    msoTable: 19,
    msoChart: 3,
};

/**
 * Auto shape type constants
 */
const MsoAutoShapeType = {
    msoShapeRectangle: 1,
    msoShapeOval: 9,
    msoShapeRoundedRectangle: 5,
    msoShapeDiamond: 4,
    msoShapeTriangle: 3,
};

/**
 * Paragraph alignment constants
 */
const PpParagraphAlignment = {
    ppAlignLeft: 1,
    ppAlignCenter: 2,
    ppAlignRight: 3,
    ppAlignJustify: 4,
};

/**
 * @class PowerPointConnector
 * @description High-level API for Microsoft PowerPoint automation
 */
class PowerPointConnector {
    constructor() {
        this.app = new COMObject('PowerPoint.Application');
    }

    /**
     * Set PowerPoint application visibility
     * @param {boolean} visible - True to show PowerPoint window
     */
    setVisible(visible) {
        try {
            // PowerPoint uses different visibility model than Word/Excel
            // Just try to set it, but don't fail if not supported
            this.app.setProperty('Visible', visible ? 1 : 0);
        } catch (e) {
            // Visibility may not be settable in some PowerPoint versions
        }
    }

    /**
     * Get PowerPoint application visibility
     * @returns {boolean} Visibility status
     */
    getVisible() {
        try {
            return this.app.getProperty('Visible');
        } catch (e) {
            return true; // Default to visible if property not accessible
        }
    }

    /**
     * Set display alerts
     * @param {boolean} display - True to show alerts
     */
    setDisplayAlerts(display) {
        try {
            this.app.setProperty('DisplayAlerts', display ? 2 : 0);
        } catch (e) {
            // DisplayAlerts may not be supported
        }
    }

    /**
     * Get PowerPoint version
     * @returns {string} PowerPoint version
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Create new presentation
     * @returns {Presentation} New presentation wrapper
     */
    addPresentation() {
        const presentations = this.app.getProperty('Presentations');
        const pres = presentations.invoke('Add', -1);
        return new Presentation(pres);
    }

    /**
     * Open existing presentation
     * @param {string} path - Full path to presentation
     * @returns {Presentation} Opened presentation wrapper
     */
    openPresentation(path) {
        const presentations = this.app.getProperty('Presentations');
        const pres = presentations.invoke('Open', path);
        return new Presentation(pres);
    }

    /**
     * Get active presentation
     * @returns {Presentation} Active presentation wrapper
     */
    getActivePresentation() {
        const pres = this.app.getProperty('ActivePresentation');
        return new Presentation(pres);
    }

    /**
     * Get presentations collection
     * @returns {COMObject} Presentations collection
     */
    getPresentations() {
        return this.app.getProperty('Presentations');
    }

    /**
     * Get count of open presentations
     * @returns {number} Number of open presentations
     */
    getPresentationCount() {
        const presentations = this.getPresentations();
        return presentations.getProperty('Count');
    }

    /**
     * Quit PowerPoint application
     */
    quit() {
        this.app.invoke('Quit');
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

/**
 * @class Presentation
 * @description Wrapper for PowerPoint Presentation
 */
class Presentation {
    constructor(comObject) {
        this.presentation = comObject;
    }

    /**
     * Get presentation name
     * @returns {string} Presentation name
     */
    getName() {
        return this.presentation.getProperty('Name');
    }

    /**
     * Get full presentation path
     * @returns {string} Full path
     */
    getFullName() {
        return this.presentation.getProperty('FullName');
    }

    /**
     * Get presentation directory path
     * @returns {string} Directory path
     */
    getPath() {
        return this.presentation.getProperty('Path');
    }

    /**
     * Get save status
     * @returns {boolean} True if saved
     */
    getSaved() {
        return this.presentation.getProperty('Saved');
    }

    /**
     * Save presentation
     */
    save() {
        this.presentation.invoke('Save');
    }

    /**
     * Save presentation with new name/format
     * @param {string} path - Full path for saved presentation
     * @param {number} format - Save format constant (PpSaveAsFileType)
     */
    saveAs(path, format = PpSaveAsFileType.ppSaveAsOpenXMLPresentation) {
        this.presentation.invoke('SaveAs', path, format);
    }

    /**
     * Close presentation
     */
    close() {
        this.presentation.invoke('Close');
    }

    /**
     * Export presentation to PDF
     * @param {string} path - Full path for PDF file
     */
    exportToPDF(path) {
        this.presentation.invoke('ExportAsFixedFormat', path, PpSaveAsFileType.ppSaveAsPDF);
    }

    /**
     * Print presentation
     */
    printOut() {
        this.presentation.invoke('PrintOut');
    }

    /**
     * Get slides collection
     * @returns {COMObject} Slides collection
     */
    getSlides() {
        return this.presentation.getProperty('Slides');
    }

    /**
     * Get slide count
     * @returns {number} Number of slides
     */
    getSlideCount() {
        const slides = this.getSlides();
        return slides.getProperty('Count');
    }

    /**
     * Get specific slide
     * @param {number} index - Slide index (1-based)
     * @returns {Slide} Slide wrapper
     */
    getSlide(index) {
        const slides = this.getSlides();
        const slide = slides.invoke('Item', index);
        return new Slide(slide);
    }

    /**
     * Add new slide
     * @param {number} index - Position to insert (1-based)
     * @param {number} layout - Slide layout constant (PpSlideLayout)
     * @returns {Slide} New slide wrapper
     */
    addSlide(index, layout = PpSlideLayout.ppLayoutBlank) {
        const slides = this.getSlides();
        const slide = slides.invoke('Add', index, layout);
        return new Slide(slide);
    }

    /**
     * Delete slide
     * @param {number} index - Slide index (1-based)
     */
    deleteSlide(index) {
        const slide = this.getSlide(index);
        slide.delete();
    }

    /**
     * Release COM object
     */
    release() {
        this.presentation.release();
    }
}

/**
 * @class Slide
 * @description Wrapper for PowerPoint Slide
 */
class Slide {
    constructor(comObject) {
        this.slide = comObject;
    }

    /**
     * Get slide index
     * @returns {number} Slide position (1-based)
     */
    getSlideIndex() {
        return this.slide.getProperty('SlideIndex');
    }

    /**
     * Get shapes collection
     * @returns {COMObject} Shapes collection
     */
    getShapes() {
        return this.slide.getProperty('Shapes');
    }

    /**
     * Get shape count
     * @returns {number} Number of shapes
     */
    getShapeCount() {
        const shapes = this.getShapes();
        return shapes.getProperty('Count');
    }

    /**
     * Get specific shape
     * @param {number} index - Shape index (1-based)
     * @returns {Shape} Shape wrapper
     */
    getShape(index) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('Item', index);
        return new Shape(shape);
    }

    /**
     * Get shape by name
     * @param {string} name - Shape name
     * @returns {Shape} Shape wrapper
     */
    getShapeByName(name) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('Item', name);
        return new Shape(shape);
    }

    /**
     * Add shape
     * @param {number} type - Shape type (MsoAutoShapeType)
     * @param {number} left - Left position
     * @param {number} top - Top position
     * @param {number} width - Width
     * @param {number} height - Height
     * @returns {Shape} New shape wrapper
     */
    addShape(type, left, top, width, height) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('AddShape', type, left, top, width, height);
        return new Shape(shape);
    }

    /**
     * Add text box
     * @param {number} left - Left position
     * @param {number} top - Top position
     * @param {number} width - Width
     * @param {number} height - Height
     * @returns {Shape} New text box wrapper
     */
    addTextBox(left, top, width, height) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('AddTextbox', 1, left, top, width, height);
        return new Shape(shape);
    }

    /**
     * Add picture
     * @param {string} filename - Full path to image file
     * @param {number} left - Left position
     * @param {number} top - Top position
     * @param {number} width - Width (optional, -1 for original)
     * @param {number} height - Height (optional, -1 for original)
     * @returns {Shape} New picture wrapper
     */
    addPicture(filename, left, top, width = -1, height = -1) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('AddPicture', filename, 0, -1, left, top, width, height);
        return new Shape(shape);
    }

    /**
     * Add table
     * @param {number} rows - Number of rows
     * @param {number} cols - Number of columns
     * @param {number} left - Left position
     * @param {number} top - Top position
     * @param {number} width - Width
     * @param {number} height - Height
     * @returns {Shape} New table wrapper
     */
    addTable(rows, cols, left, top, width, height) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('AddTable', rows, cols, left, top, width, height);
        return new Shape(shape);
    }

    /**
     * Delete slide
     */
    delete() {
        this.slide.invoke('Delete');
    }

    /**
     * Copy slide
     */
    copy() {
        this.slide.invoke('Copy');
    }

    /**
     * Move slide to position
     * @param {number} toPos - Target position (1-based)
     */
    moveTo(toPos) {
        this.slide.invoke('MoveTo', toPos);
    }

    /**
     * Get notes text
     * @returns {string} Notes text
     * @experimental This method may have COM limitations in some PowerPoint versions
     */
    getNotes() {
        try {
            const notesPage = this.slide.getProperty('NotesPage');
            const shapes = notesPage.getProperty('Shapes');
            const placeholders = shapes.invoke('Placeholders');
            const notesShape = placeholders.invoke('Item', 2);
            const textFrame = notesShape.getProperty('TextFrame');
            const textRange = textFrame.getProperty('TextRange');
            return textRange.getProperty('Text');
        } catch (e) {
            return '';
        }
    }

    /**
     * Set notes text
     * @param {string} text - Notes text
     * @experimental This method may have COM limitations in some PowerPoint versions
     */
    setNotes(text) {
        try {
            // Method 1: Use Placeholders
            const notesPage = this.slide.getProperty('NotesPage');
            const shapes = notesPage.getProperty('Shapes');
            const placeholders = shapes.invoke('Placeholders');
            const notesShape = placeholders.invoke('Item', 2);
            const textFrame = notesShape.getProperty('TextFrame');
            const textRange = textFrame.getProperty('TextRange');
            textRange.setProperty('Text', text);
        } catch (e1) {
            try {
                // Method 2: Iterate through shapes to find notes placeholder
                const notesPage = this.slide.getProperty('NotesPage');
                const shapes = notesPage.getProperty('Shapes');
                const count = shapes.getProperty('Count');
                
                for (let i = 1; i <= count; i++) {
                    const shape = shapes.invoke('Item', i);
                    const shapeType = shape.getProperty('Type');
                    if (shapeType === 14) { // msoPlaceholder
                        try {
                            const placeholderFormat = shape.getProperty('PlaceholderFormat');
                            const placeholderType = placeholderFormat.getProperty('Type');
                            if (placeholderType === 2) { // ppPlaceholderBody (notes)
                                const textFrame = shape.getProperty('TextFrame');
                                const textRange = textFrame.getProperty('TextRange');
                                textRange.setProperty('Text', text);
                                return;
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                }
                throw new Error('Notes placeholder not found');
            } catch (e2) {
                try {
                    // Method 3: Direct access to second shape (common pattern)
                    const notesPage = this.slide.getProperty('NotesPage');
                    const shapes = notesPage.getProperty('Shapes');
                    const notesShape = shapes.invoke('Item', 2);
                    const textFrame = notesShape.getProperty('TextFrame');
                    const textRange = textFrame.getProperty('TextRange');
                    textRange.setProperty('Text', text);
                } catch (e3) {
                    throw new Error('Failed to set notes: ' + e3.message);
                }
            }
        }
    }

    /**
     * Set slide transition
     * @param {number} type - Transition effect type (PpEntryEffect)
     * @param {number} speed - Transition speed (1=Fast, 2=Medium, 3=Slow)
     */
    setTransition(type, speed = 2) {
        const slideShowTransition = this.slide.getProperty('SlideShowTransition');
        slideShowTransition.setProperty('EntryEffect', type);
        slideShowTransition.setProperty('Speed', speed);
    }
    
    /**
     * Set transition duration
     * @param {number} seconds - Duration in seconds
     */
    setTransitionDuration(seconds) {
        const slideShowTransition = this.slide.getProperty('SlideShowTransition');
        slideShowTransition.setProperty('Duration', seconds);
    }
    
    /**
     * Set transition advance on time
     * @param {boolean} advanceOnTime - True to advance automatically
     * @param {number} advanceTime - Time in seconds
     */
    setTransitionAdvanceOnTime(advanceOnTime, advanceTime = 0) {
        const slideShowTransition = this.slide.getProperty('SlideShowTransition');
        slideShowTransition.setProperty('AdvanceOnTime', advanceOnTime ? -1 : 0);
        if (advanceOnTime) {
            slideShowTransition.setProperty('AdvanceTime', advanceTime);
        }
    }
    
    /**
     * Set transition advance on click
     * @param {boolean} advanceOnClick - True to advance on click
     */
    setTransitionAdvanceOnClick(advanceOnClick) {
        const slideShowTransition = this.slide.getProperty('SlideShowTransition');
        slideShowTransition.setProperty('AdvanceOnClick', advanceOnClick ? -1 : 0);
    }

    /**
     * Release COM object
     */
    release() {
        this.slide.release();
    }
}

/**
 * @class Shape
 * @description Wrapper for PowerPoint Shape
 */
class Shape {
    constructor(comObject) {
        this.shape = comObject;
    }

    /**
     * Get shape name
     * @returns {string} Shape name
     */
    getName() {
        return this.shape.getProperty('Name');
    }

    /**
     * Set shape name
     * @param {string} name - Shape name
     */
    setName(name) {
        this.shape.setProperty('Name', name);
    }

    /**
     * Get shape type
     * @returns {number} Shape type constant
     */
    getType() {
        return this.shape.getProperty('Type');
    }

    /**
     * Get left position
     * @returns {number} Left position in points
     */
    getLeft() {
        return this.shape.getProperty('Left');
    }

    /**
     * Set left position
     * @param {number} value - Left position in points
     */
    setLeft(value) {
        this.shape.setProperty('Left', value);
    }

    /**
     * Get top position
     * @returns {number} Top position in points
     */
    getTop() {
        return this.shape.getProperty('Top');
    }

    /**
     * Set top position
     * @param {number} value - Top position in points
     */
    setTop(value) {
        this.shape.setProperty('Top', value);
    }

    /**
     * Get width
     * @returns {number} Width in points
     */
    getWidth() {
        return this.shape.getProperty('Width');
    }

    /**
     * Set width
     * @param {number} value - Width in points
     */
    setWidth(value) {
        this.shape.setProperty('Width', value);
    }

    /**
     * Get height
     * @returns {number} Height in points
     */
    getHeight() {
        return this.shape.getProperty('Height');
    }

    /**
     * Set height
     * @param {number} value - Height in points
     */
    setHeight(value) {
        this.shape.setProperty('Height', value);
    }

    /**
     * Check if shape has text frame
     * @returns {boolean} True if has text frame
     */
    hasTextFrame() {
        try {
            const textFrame = this.shape.getProperty('TextFrame');
            return textFrame !== null;
        } catch (e) {
            return false;
        }
    }

    /**
     * Get text frame
     * @returns {TextFrame} Text frame wrapper
     */
    getTextFrame() {
        const textFrame = this.shape.getProperty('TextFrame');
        return new TextFrame(textFrame);
    }

    /**
     * Get text content
     * @returns {string} Text content
     */
    getText() {
        if (this.hasTextFrame()) {
            const textFrame = this.getTextFrame();
            return textFrame.getText();
        }
        return '';
    }

    /**
     * Set text content
     * @param {string} text - Text to set
     */
    setText(text) {
        if (this.hasTextFrame()) {
            const textFrame = this.getTextFrame();
            textFrame.setText(text);
        }
    }

    /**
     * Check if shape is table
     * @returns {boolean} True if table
     */
    hasTable() {
        try {
            const table = this.shape.getProperty('Table');
            return table !== null;
        } catch (e) {
            return false;
        }
    }

    /**
     * Get table
     * @returns {PPTable} Table wrapper
     */
    getTable() {
        const table = this.shape.getProperty('Table');
        return new PPTable(table);
    }

    /**
     * Delete shape
     */
    delete() {
        this.shape.invoke('Delete');
    }

    /**
     * Add animation effect to shape
     * @param {number} effect - Animation effect type (msoAnimEffect constants)
     * @param {number} trigger - Trigger type (0=OnClick, 1=WithPrevious, 2=AfterPrevious)
     * @returns {Object} Animation effect object
     */
    addAnimation(effect, trigger = 0) {
        try {
            const slide = this.shape.getProperty('Parent');
            const timeline = slide.getProperty('TimeLine');
            const mainSequence = timeline.getProperty('MainSequence');
            const animEffect = mainSequence.invoke('AddEffect', this.shape, effect, 0, trigger);
            return animEffect;
        } catch (e) {
            throw new Error('Failed to add animation: ' + e.message);
        }
    }
    
    /**
     * Remove all animations from shape
     */
    removeAnimations() {
        try {
            const slide = this.shape.getProperty('Parent');
            const timeline = slide.getProperty('TimeLine');
            const mainSequence = timeline.getProperty('MainSequence');
            const effects = mainSequence.getProperty('Effects');
            const count = effects.getProperty('Count');
            
            for (let i = count; i >= 1; i--) {
                const effect = effects.invoke('Item', i);
                const effectShape = effect.getProperty('Shape');
                if (effectShape.getProperty('Name') === this.getName()) {
                    effect.invoke('Delete');
                }
            }
        } catch (e) {
            // Ignore errors if no animations exist
        }
    }

    /**
     * Copy shape
     */
    copy() {
        this.shape.invoke('Copy');
    }

    /**
     * Release COM object
     */
    release() {
        this.shape.release();
    }
}

/**
 * @class TextFrame
 * @description Wrapper for PowerPoint TextFrame
 */
class TextFrame {
    constructor(comObject) {
        this.textFrame = comObject;
    }

    /**
     * Get text range
     * @returns {TextRange} Text range wrapper
     */
    getTextRange() {
        const textRange = this.textFrame.getProperty('TextRange');
        return new TextRange(textRange);
    }

    /**
     * Get text content
     * @returns {string} Text content
     */
    getText() {
        const textRange = this.getTextRange();
        return textRange.getText();
    }

    /**
     * Set text content
     * @param {string} text - Text to set
     */
    setText(text) {
        const textRange = this.getTextRange();
        textRange.setText(text);
    }

    /**
     * Release COM object
     */
    release() {
        this.textFrame.release();
    }
}

/**
 * @class TextRange
 * @description Wrapper for PowerPoint TextRange
 */
class TextRange {
    constructor(comObject) {
        this.textRange = comObject;
    }

    /**
     * Get text content
     * @returns {string} Text content
     */
    getText() {
        return this.textRange.getProperty('Text');
    }

    /**
     * Set text content
     * @param {string} text - Text to set
     */
    setText(text) {
        this.textRange.setProperty('Text', text);
    }

    /**
     * Get font
     * @returns {PPFont} Font wrapper
     */
    getFont() {
        const font = this.textRange.getProperty('Font');
        return new PPFont(font);
    }

    /**
     * Get paragraph format
     * @returns {ParagraphFormat} Paragraph format wrapper
     */
    getParagraphFormat() {
        const format = this.textRange.getProperty('ParagraphFormat');
        return new ParagraphFormat(format);
    }

    /**
     * Release COM object
     */
    release() {
        this.textRange.release();
    }
}

/**
 * @class PPFont
 * @description Wrapper for PowerPoint Font
 */
class PPFont {
    constructor(comObject) {
        this.font = comObject;
    }

    /**
     * Get font name
     * @returns {string} Font name
     */
    getName() {
        return this.font.getProperty('Name');
    }

    /**
     * Set font name
     * @param {string} value - Font name
     */
    setName(value) {
        this.font.setProperty('Name', value);
    }

    /**
     * Get font size
     * @returns {number} Font size in points
     */
    getSize() {
        return this.font.getProperty('Size');
    }

    /**
     * Set font size
     * @param {number} value - Font size in points
     */
    setSize(value) {
        this.font.setProperty('Size', value);
    }

    /**
     * Get bold
     * @returns {boolean} True if bold
     */
    getBold() {
        return this.font.getProperty('Bold');
    }

    /**
     * Set bold
     * @param {boolean} value - True for bold
     */
    setBold(value) {
        this.font.setProperty('Bold', value ? -1 : 0);
    }

    /**
     * Get italic
     * @returns {boolean} True if italic
     */
    getItalic() {
        return this.font.getProperty('Italic');
    }

    /**
     * Set italic
     * @param {boolean} value - True for italic
     */
    setItalic(value) {
        this.font.setProperty('Italic', value ? -1 : 0);
    }

    /**
     * Get underline
     * @returns {boolean} True if underline
     */
    getUnderline() {
        return this.font.getProperty('Underline');
    }

    /**
     * Set underline
     * @param {boolean} value - True for underline
     */
    setUnderline(value) {
        this.font.setProperty('Underline', value ? -1 : 0);
    }

    /**
     * Get color
     * @returns {number} Color value (RGB)
     */
    getColor() {
        return this.font.getProperty('Color').getProperty('RGB');
    }

    /**
     * Set color
     * @param {number} value - Color value (RGB)
     */
    setColor(value) {
        this.font.getProperty('Color').setProperty('RGB', value);
    }

    /**
     * Release COM object
     */
    release() {
        this.font.release();
    }
}

/**
 * @class ParagraphFormat
 * @description Wrapper for PowerPoint ParagraphFormat
 */
class ParagraphFormat {
    constructor(comObject) {
        this.format = comObject;
    }

    /**
     * Get alignment
     * @returns {number} Alignment constant
     */
    getAlignment() {
        return this.format.getProperty('Alignment');
    }

    /**
     * Set alignment
     * @param {number} value - Alignment constant (PpParagraphAlignment)
     */
    setAlignment(value) {
        this.format.setProperty('Alignment', value);
    }

    /**
     * Release COM object
     */
    release() {
        this.format.release();
    }
}

/**
 * @class PPTable
 * @description Wrapper for PowerPoint Table
 */
class PPTable {
    constructor(comObject) {
        this.table = comObject;
    }

    /**
     * Get rows collection
     * @returns {COMObject} Rows collection
     */
    getRows() {
        return this.table.getProperty('Rows');
    }

    /**
     * Get columns collection
     * @returns {COMObject} Columns collection
     */
    getColumns() {
        return this.table.getProperty('Columns');
    }

    /**
     * Get row count
     * @returns {number} Number of rows
     */
    getRowCount() {
        const rows = this.getRows();
        return rows.getProperty('Count');
    }

    /**
     * Get column count
     * @returns {number} Number of columns
     */
    getColumnCount() {
        const columns = this.getColumns();
        return columns.getProperty('Count');
    }

    /**
     * Get cell
     * @param {number} row - Row index (1-based)
     * @param {number} col - Column index (1-based)
     * @returns {COMObject} Cell object
     */
    getCell(row, col) {
        return this.table.invoke('Cell', row, col);
    }

    /**
     * Get cell text
     * @param {number} row - Row index (1-based)
     * @param {number} col - Column index (1-based)
     * @returns {string} Cell text
     */
    getCellText(row, col) {
        const cell = this.getCell(row, col);
        const shape = cell.getProperty('Shape');
        const textFrame = shape.getProperty('TextFrame');
        const textRange = textFrame.getProperty('TextRange');
        return textRange.getProperty('Text');
    }

    /**
     * Set cell text
     * @param {number} row - Row index (1-based)
     * @param {number} col - Column index (1-based)
     * @param {string} text - Text to set
     */
    setCellText(row, col, text) {
        const cell = this.getCell(row, col);
        const shape = cell.getProperty('Shape');
        const textFrame = shape.getProperty('TextFrame');
        const textRange = textFrame.getProperty('TextRange');
        textRange.setProperty('Text', text);
    }

    /**
     * Release COM object
     */
    release() {
        this.table.release();
    }
}

module.exports = {
    PowerPointConnector,
    Presentation,
    Slide,
    Shape,
    TextFrame,
    TextRange,
    PPFont,
    ParagraphFormat,
    PPTable,
    PpSaveAsFileType,
    PpSlideLayout,
    MsoShapeType,
    MsoAutoShapeType,
    PpParagraphAlignment,
};
