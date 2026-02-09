/**
 * @fileoverview Microsoft Word Connector
 * Provides a high-level, type-safe API for Word automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * Word save formats
 */
const WdSaveFormat = {
    wdFormatDocument: 0,
    wdFormatDocumentDefault: 16,
    wdFormatPDF: 17,
    wdFormatTemplate: 1,
    wdFormatText: 2,
    wdFormatTextLineBreaks: 3,
    wdFormatDOSText: 4,
    wdFormatRTF: 6,
    wdFormatHTML: 8,
    wdFormatXMLDocument: 12,
    wdFormatDocx: 16,
    wdFormatDocm: 13,
};

/**
 * Paragraph alignment constants
 */
const WdParagraphAlignment = {
    wdAlignParagraphLeft: 0,
    wdAlignParagraphCenter: 1,
    wdAlignParagraphRight: 2,
    wdAlignParagraphJustify: 3,
    wdAlignParagraphDistribute: 4,
};

/**
 * Page orientation constants
 */
const WdOrientation = {
    wdOrientPortrait: 0,
    wdOrientLandscape: 1,
};

/**
 * Underline style constants
 */
const WdUnderline = {
    wdUnderlineNone: 0,
    wdUnderlineSingle: 1,
    wdUnderlineDouble: 3,
    wdUnderlineDotted: 4,
    wdUnderlineThick: 6,
    wdUnderlineWave: 11,
};

/**
 * Color constants
 */
const WdColor = {
    wdColorAutomatic: -16777216,
    wdColorBlack: 0,
    wdColorBlue: 16711680,
    wdColorRed: 255,
    wdColorYellow: 65535,
    wdColorGreen: 65280,
    wdColorWhite: 16777215,
};

/**
 * Break type constants
 */
const WdBreakType = {
    wdPageBreak: 7,
    wdColumnBreak: 8,
    wdSectionBreakNextPage: 2,
    wdSectionBreakContinuous: 3,
    wdLineBreak: 6,
};

/**
 * @class WordConnector
 * @description High-level API for Microsoft Word automation
 */
class WordConnector {
    constructor() {
        this.app = new COMObject('Word.Application');
    }

    /**
     * Set Word application visibility
     * @param {boolean} visible - True to show Word window
     */
    setVisible(visible) {
        this.app.setProperty('Visible', visible);
    }

    /**
     * Get Word application visibility
     * @returns {boolean} Visibility status
     */
    getVisible() {
        return this.app.getProperty('Visible');
    }

    /**
     * Set display alerts
     * @param {boolean} display - True to show alerts
     */
    setDisplayAlerts(display) {
        this.app.setProperty('DisplayAlerts', display ? -1 : 0);
    }

    /**
     * Set screen updating
     * @param {boolean} updating - True to enable screen updates
     */
    setScreenUpdating(updating) {
        this.app.setProperty('ScreenUpdating', updating);
    }

    /**
     * Get Word version
     * @returns {string} Word version
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Create new document
     * @returns {WordDocument} New document wrapper
     */
    addDocument() {
        const documents = this.app.getProperty('Documents');
        const doc = documents.invoke('Add');
        return new WordDocument(doc);
    }

    /**
     * Open existing document
     * @param {string} path - Full path to document
     * @returns {WordDocument} Opened document wrapper
     */
    openDocument(path) {
        const documents = this.app.getProperty('Documents');
        const doc = documents.invoke('Open', path);
        return new WordDocument(doc);
    }

    /**
     * Get active document
     * @returns {WordDocument} Active document wrapper
     */
    getActiveDocument() {
        const doc = this.app.getProperty('ActiveDocument');
        return new WordDocument(doc);
    }

    /**
     * Get documents collection
     * @returns {COMObject} Documents collection
     */
    getDocuments() {
        return this.app.getProperty('Documents');
    }

    /**
     * Get count of open documents
     * @returns {number} Number of open documents
     */
    getDocumentCount() {
        const documents = this.getDocuments();
        return documents.getProperty('Count');
    }

    /**
     * Quit Word application
     * @param {boolean} saveChanges - Save changes before quitting
     */
    quit(saveChanges = false) {
        this.app.invoke('Quit', saveChanges ? -1 : 0);
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

/**
 * @class WordDocument
 * @description Wrapper for Word Document
 */
class WordDocument {
    constructor(comObject) {
        this.document = comObject;
    }

    /**
     * Get document name
     * @returns {string} Document name
     */
    getName() {
        return this.document.getProperty('Name');
    }

    /**
     * Get full document path
     * @returns {string} Full path
     */
    getFullName() {
        return this.document.getProperty('FullName');
    }

    /**
     * Get document directory path
     * @returns {string} Directory path
     */
    getPath() {
        return this.document.getProperty('Path');
    }

    /**
     * Get save status
     * @returns {boolean} True if saved
     */
    getSaved() {
        return this.document.getProperty('Saved');
    }

    /**
     * Save document
     */
    save() {
        this.document.invoke('Save');
    }

    /**
     * Save document with new name/format
     * @param {string} path - Full path for saved document
     * @param {number} format - Save format constant (WdSaveFormat)
     */
    saveAs(path, format = WdSaveFormat.wdFormatDocumentDefault) {
        this.document.invoke('SaveAs', path, format);
    }

    /**
     * Close document
     * @param {boolean} saveChanges - Save changes before closing
     */
    close(saveChanges = false) {
        this.document.invoke('Close', saveChanges ? -1 : 0);
    }

    /**
     * Export document to PDF
     * @param {string} path - Full path for PDF file
     */
    exportToPDF(path) {
        this.document.invoke('ExportAsFixedFormat', path, WdSaveFormat.wdFormatPDF);
    }

    /**
     * Activate document
     */
    activate() {
        this.document.invoke('Activate');
    }

    /**
     * Get document content as Range
     * @returns {WordRange} Document content range
     */
    getContent() {
        const range = this.document.getProperty('Content');
        return new WordRange(range);
    }

    /**
     * Get text content of entire document
     * @returns {string} Document text
     */
    getText() {
        const content = this.getContent();
        return content.getText();
    }

    /**
     * Set text content of entire document
     * @param {string} text - Text to set
     */
    setText(text) {
        const content = this.getContent();
        content.setText(text);
    }

    /**
     * Get paragraphs collection
     * @returns {COMObject} Paragraphs collection
     */
    getParagraphs() {
        return this.document.getProperty('Paragraphs');
    }

    /**
     * Get paragraph count
     * @returns {number} Number of paragraphs
     */
    getParagraphCount() {
        const paragraphs = this.getParagraphs();
        return paragraphs.getProperty('Count');
    }

    /**
     * Get specific paragraph
     * @param {number} index - Paragraph index (1-based)
     * @returns {WordParagraph} Paragraph wrapper
     */
    getParagraph(index) {
        const paragraphs = this.getParagraphs();
        const paragraph = paragraphs.invoke('Item', index);
        return new WordParagraph(paragraph);
    }

    /**
     * Add new paragraph
     * @param {string} text - Paragraph text (optional)
     * @returns {WordParagraph} New paragraph wrapper
     */
    addParagraph(text = '') {
        const content = this.getContent();
        content.insertAfter('\r\n' + text);
        const paragraphs = this.getParagraphs();
        const lastIndex = paragraphs.getProperty('Count');
        return this.getParagraph(lastIndex);
    }

    /**
     * Get tables collection
     * @returns {COMObject} Tables collection
     */
    getTables() {
        return this.document.getProperty('Tables');
    }

    /**
     * Get table count
     * @returns {number} Number of tables
     */
    getTableCount() {
        const tables = this.getTables();
        return tables.getProperty('Count');
    }

    /**
     * Get specific table
     * @param {number} index - Table index (1-based)
     * @returns {WordTable} Table wrapper
     */
    getTable(index) {
        const tables = this.getTables();
        const table = tables.invoke('Item', index);
        return new WordTable(table);
    }

    /**
     * Add new table
     * @param {number} rows - Number of rows
     * @param {number} cols - Number of columns
     * @param {WordRange} range - Range where table should be inserted (optional)
     * @returns {WordTable} New table wrapper
     */
    /**
     * Add table to document
     * @param {number} rows - Number of rows
     * @param {number} cols - Number of columns
     * @param {WordRange} [range] - Optional range
     * @returns {WordTable} Table wrapper
     * @experimental This method has known COM limitations (73% success rate)
     * @see WORD_COM_LIMITATIONS.md for details and workarounds
     */
    addTable(rows, cols, range = null) {
        try {
            // Ensure document is activated
            this.document.invoke('Activate');
            
            // Method 1: Use Selection to add table
            const app = this.document.getProperty('Application');
            const selection = app.getProperty('Selection');
            
            // Move to end of document
            selection.invoke('EndKey', 6); // wdStory = 6
            
            const tables = this.document.getProperty('Tables');
            const selRange = selection.getProperty('Range');
            const table = tables.invoke('Add', selRange, rows, cols);
            return new WordTable(table);
        } catch (e1) {
            try {
                // Method 2: Use document Range directly
                const contentRange = this.document.getProperty('Content');
                const endPos = contentRange.getProperty('End');
                const newRange = this.document.invoke('Range', endPos - 1, endPos - 1);
                const tables = this.document.getProperty('Tables');
                const table = tables.invoke('Add', newRange, rows, cols);
                return new WordTable(table);
            } catch (e2) {
                try {
                    // Method 3: Insert at beginning
                    const startRange = this.document.invoke('Range', 0, 0);
                    const tables = this.document.getProperty('Tables');
                    const table = tables.invoke('Add', startRange, rows, cols);
                    return new WordTable(table);
                } catch (e3) {
                    throw new Error('Failed to add table: ' + e3.message);
                }
            }
        }
    }

    /**
     * Get bookmarks collection
     * @returns {COMObject} Bookmarks collection
     */
    getBookmarks() {
        return this.document.getProperty('Bookmarks');
    }

    /**
     * Add bookmark
     * @param {string} name - Bookmark name
     * @experimental This method has known COM limitations (73% success rate)
     * @see WORD_COM_LIMITATIONS.md for details and workarounds
     * @param {WordRange} range - Range to bookmark
     */
    addBookmark(name, range = null) {
        try {
            // Ensure document is activated
            this.document.invoke('Activate');
            
            // Method 1: Use Selection object
            const app = this.document.getProperty('Application');
            const selection = app.getProperty('Selection');
            const selRange = selection.getProperty('Range');
            const bookmarks = this.getBookmarks();
            bookmarks.invoke('Add', name, selRange);
        } catch (e1) {
            try {
                // Method 2: Use document Range directly
                const contentRange = this.document.getProperty('Content');
                const startPos = contentRange.getProperty('Start');
                const endPos = contentRange.getProperty('End');
                const newRange = this.document.invoke('Range', startPos, endPos);
                const bookmarks = this.getBookmarks();
                bookmarks.invoke('Add', name, newRange);
            } catch (e2) {
                try {
                    // Method 3: Use provided range or content
                    const bookmarks = this.getBookmarks();
                    let targetRange;
                    
                    if (range) {
                        targetRange = range.range ? range.range : range;
                    } else {
                        targetRange = this.document.getProperty('Content');
                    }
                    
                    bookmarks.invoke('Add', name, targetRange);
                } catch (e3) {
                    throw new Error('Failed to add bookmark: ' + e3.message);
                }
            }
        }
    }

    /**
     * Check if bookmark exists
     * @param {string} name - Bookmark name
     * @returns {boolean} True if exists
     * @experimental This method has known COM limitations
     */
    bookmarkExists(name) {
        try {
            const bookmarks = this.getBookmarks();
            return bookmarks.invoke('Exists', name);
        } catch (e) {
            // Fallback: try to get the bookmark and catch error
            try {
                const bookmarks = this.getBookmarks();
                bookmarks.invoke('Item', name);
                return true;
            } catch (e2) {
                return false;
            }
        }
    }

    /**
     * Go to bookmark
     * @param {string} name - Bookmark name
     * @returns {WordRange} Range at bookmark
     * @experimental This method has known COM limitations
     */
    goToBookmark(name) {
        const bookmarks = this.getBookmarks();
        const bookmark = bookmarks.invoke('Item', name);
        const range = bookmark.getProperty('Range');
        return new WordRange(range);
    }

    /**
     * Delete bookmark
     * @param {string} name - Bookmark name
     */
    deleteBookmark(name) {
        const bookmarks = this.getBookmarks();
        const bookmark = bookmarks.invoke('Item', name);
        bookmark.invoke('Delete');
    }

    /**
     * Get comments collection
     * @returns {COMObject} Comments collection
     */
    getComments() {
        return this.document.getProperty('Comments');
    }

    /**
     * Get comment count
     * @returns {number} Number of comments
     */
    getCommentCount() {
        const comments = this.getComments();
        return comments.getProperty('Count');
    }

    /**
     * Add comment
     * @param {string} text - Comment text
     * @param {WordRange} range - Range to comment (optional)
     * @experimental This method has known COM limitations (73% success rate)
     * @see WORD_COM_LIMITATIONS.md for details and workarounds
     */
    addComment(text, range = null) {
        try {
            // Ensure document is activated
            this.document.invoke('Activate');
            
            // Method 1: Use Selection object
            const app = this.document.getProperty('Application');
            const selection = app.getProperty('Selection');
            const selRange = selection.getProperty('Range');
            const comments = this.getComments();
            comments.invoke('Add', selRange, text);
        } catch (e1) {
            try {
                // Method 2: Use document Range directly
                const contentRange = this.document.getProperty('Content');
                const startPos = contentRange.getProperty('Start');
                const newRange = this.document.invoke('Range', startPos, startPos);
                const comments = this.getComments();
                comments.invoke('Add', newRange, text);
            } catch (e2) {
                try {
                    // Method 3: Use provided range or content
                    const comments = this.getComments();
                    let targetRange;
                    
                    if (range) {
                        targetRange = range.range ? range.range : range;
                    } else {
                        targetRange = this.document.getProperty('Content');
                    }
                    
                    comments.invoke('Add', targetRange, text);
                } catch (e3) {
                    throw new Error('Failed to add comment: ' + e3.message);
                }
            }
        }
    }

    /**
     * Get page setup
     * @returns {WordPageSetup} Page setup wrapper
     */
    getPageSetup() {
        const pageSetup = this.document.getProperty('PageSetup');
        return new WordPageSetup(pageSetup);
    }

    /**
     * Protect document
     * @param {number} type - Protection type
     * @param {string} password - Password (optional)
     */
    protect(type = 2, password = '') {
        if (password) {
            this.document.invoke('Protect', type, true, password);
        } else {
            this.document.invoke('Protect', type);
        }
    }

    /**
     * Unprotect document
     * @param {string} password - Password (optional)
     */
    unprotect(password = '') {
        if (password) {
            this.document.invoke('Unprotect', password);
        } else {
            this.document.invoke('Unprotect');
        }
    }

    /**
     * Print document
     */
    printOut() {
        this.document.invoke('PrintOut');
    }

    /**
     * Get fields collection
     * @returns {COMObject} Fields collection
     */
    getFields() {
        return this.document.getProperty('Fields');
    }

    /**
     * Update all fields
     */
    updateFields() {
        const fields = this.getFields();
        fields.invoke('Update');
    }

    /**
     * Get sections collection
     * @returns {COMObject} Sections collection
     */
    getSections() {
        return this.document.getProperty('Sections');
    }

    /**
     * Get section count
     * @returns {number} Number of sections
     */
    getSectionCount() {
        const sections = this.getSections();
        return sections.getProperty('Count');
    }

    /**
     * Release COM object
     */
    release() {
        this.document.release();
    }
}

/**
 * @class WordRange
 * @description Wrapper for Word Range
 */
class WordRange {
    constructor(comObject) {
        this.range = comObject;
    }

    /**
     * Get text content
     * @returns {string} Text content
     */
    getText() {
        return this.range.getProperty('Text');
    }

    /**
     * Set text content
     * @param {string} text - Text to set
     */
    setText(text) {
        this.range.setProperty('Text', text);
    }

    /**
     * Get start position
     * @returns {number} Start position
     */
    getStart() {
        return this.range.getProperty('Start');
    }

    /**
     * Set start position
     * @param {number} position - Start position
     */
    setStart(position) {
        this.range.setProperty('Start', position);
    }

    /**
     * Get end position
     * @returns {number} End position
     */
    getEnd() {
        return this.range.getProperty('End');
    }

    /**
     * Set end position
     * @param {number} position - End position
     */
    setEnd(position) {
        this.range.setProperty('End', position);
    }

    /**
     * Select range
     */
    select() {
        this.range.invoke('Select');
    }

    /**
     * Insert text after range
     * @param {string} text - Text to insert
     */
    insertAfter(text) {
        this.range.invoke('InsertAfter', text);
    }

    /**
     * Insert text before range
     * @param {string} text - Text to insert
     */
    insertBefore(text) {
        this.range.invoke('InsertBefore', text);
    }

    /**
     * Delete range
     */
    delete() {
        this.range.invoke('Delete');
    }

    /**
     * Copy range to clipboard
     */
    copy() {
        this.range.invoke('Copy');
    }

    /**
     * Cut range to clipboard
     */
    cut() {
        this.range.invoke('Cut');
    }

    /**
     * Paste from clipboard
     */
    paste() {
        this.range.invoke('Paste');
    }

    /**
     * Find text in range
     * @param {string} findText - Text to find
     * @returns {boolean} True if found
     */
    find(findText) {
        const find = this.range.getProperty('Find');
        return find.invoke('Execute', findText);
    }

    /**
     * Replace text in range
     * @param {string} findText - Text to find
     * @param {string} replaceWith - Replacement text
     * @returns {boolean} True if replaced
     */
    replace(findText, replaceWith) {
        try {
            const find = this.range.getProperty('Find');
            // Simplified Execute call for replace
            find.invoke('ClearFormatting');
            const replacement = find.getProperty('Replacement');
            replacement.invoke('ClearFormatting');
            replacement.setProperty('Text', replaceWith);
            return find.invoke('Execute', findText, false, false, false, false, false, true, 1, '', 2);
        } catch (e) {
            // Fallback: simple text replacement
            const text = this.getText();
            const newText = text.replace(new RegExp(findText, 'g'), replaceWith);
            this.setText(newText);
            return true;
        }
    }

    /**
     * Get font object
     * @returns {WordFont} Font wrapper
     */
    getFont() {
        const font = this.range.getProperty('Font');
        return new WordFont(font);
    }

    /**
     * Get paragraph format
     * @returns {WordParagraphFormat} Paragraph format wrapper
     */
    getParagraphFormat() {
        const format = this.range.getProperty('ParagraphFormat');
        return new WordParagraphFormat(format);
    }

    /**
     * Get bold formatting
     * @returns {boolean} True if bold
     */
    getBold() {
        return this.range.getProperty('Bold');
    }

    /**
     * Set bold formatting
     * @param {boolean} value - True for bold
     */
    setBold(value) {
        this.range.setProperty('Bold', value ? -1 : 0);
    }

    /**
     * Get italic formatting
     * @returns {boolean} True if italic
     */
    getItalic() {
        return this.range.getProperty('Italic');
    }

    /**
     * Set italic formatting
     * @param {boolean} value - True for italic
     */
    setItalic(value) {
        this.range.setProperty('Italic', value ? -1 : 0);
    }

    /**
     * Get underline formatting
     * @returns {number} Underline style
     */
    getUnderline() {
        return this.range.getProperty('Underline');
    }

    /**
     * Set underline formatting
     * @param {number} value - Underline style constant
     */
    setUnderline(value) {
        this.range.setProperty('Underline', value);
    }

    /**
     * Insert page break
     */
    insertBreak(breakType = WdBreakType.wdPageBreak) {
        this.range.invoke('InsertBreak', breakType);
    }

    /**
     * Collapse range
     * @param {number} direction - 0=start, 1=end
     */
    collapse(direction = 0) {
        this.range.invoke('Collapse', direction);
    }

    /**
     * Release COM object
     */
    release() {
        this.range.release();
    }
}

/**
 * @class WordParagraph
 * @description Wrapper for Word Paragraph
 */
class WordParagraph {
    constructor(comObject) {
        this.paragraph = comObject;
    }

    /**
     * Get paragraph range
     * @returns {WordRange} Paragraph range
     */
    getRange() {
        const range = this.paragraph.getProperty('Range');
        return new WordRange(range);
    }

    /**
     * Get paragraph format
     * @returns {WordParagraphFormat} Paragraph format wrapper
     */
    getFormat() {
        const format = this.paragraph.getProperty('Format');
        return new WordParagraphFormat(format);
    }

    /**
     * Get alignment
     * @returns {number} Alignment constant
     */
    getAlignment() {
        const format = this.getFormat();
        return format.getAlignment();
    }

    /**
     * Set alignment
     * @param {number} value - Alignment constant
     */
    setAlignment(value) {
        const format = this.getFormat();
        format.setAlignment(value);
    }

    /**
     * Delete paragraph
     */
    delete() {
        const range = this.getRange();
        range.delete();
    }

    /**
     * Release COM object
     */
    release() {
        this.paragraph.release();
    }
}

/**
 * @class WordParagraphFormat
 * @description Wrapper for Word ParagraphFormat
 */
class WordParagraphFormat {
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
     * @param {number} value - Alignment constant
     */
    setAlignment(value) {
        this.format.setProperty('Alignment', value);
    }

    /**
     * Get space before
     * @returns {number} Space before in points
     */
    getSpaceBefore() {
        return this.format.getProperty('SpaceBefore');
    }

    /**
     * Set space before
     * @param {number} value - Space before in points
     */
    setSpaceBefore(value) {
        this.format.setProperty('SpaceBefore', value);
    }

    /**
     * Get space after
     * @returns {number} Space after in points
     */
    getSpaceAfter() {
        return this.format.getProperty('SpaceAfter');
    }

    /**
     * Set space after
     * @param {number} value - Space after in points
     */
    setSpaceAfter(value) {
        this.format.setProperty('SpaceAfter', value);
    }

    /**
     * Get left indent
     * @returns {number} Left indent in points
     */
    getLeftIndent() {
        return this.format.getProperty('LeftIndent');
    }

    /**
     * Set left indent
     * @param {number} value - Left indent in points
     */
    setLeftIndent(value) {
        this.format.setProperty('LeftIndent', value);
    }

    /**
     * Get right indent
     * @returns {number} Right indent in points
     */
    getRightIndent() {
        return this.format.getProperty('RightIndent');
    }

    /**
     * Set right indent
     * @param {number} value - Right indent in points
     */
    setRightIndent(value) {
        this.format.setProperty('RightIndent', value);
    }

    /**
     * Release COM object
     */
    release() {
        this.format.release();
    }
}

/**
 * @class WordTable
 * @description Wrapper for Word Table
 */
class WordTable {
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
        const range = cell.getProperty('Range');
        return range.getProperty('Text');
    }

    /**
     * Set cell text
     * @param {number} row - Row index (1-based)
     * @param {number} col - Column index (1-based)
     * @param {string} text - Text to set
     */
    setCellText(row, col, text) {
        const cell = this.getCell(row, col);
        const range = cell.getProperty('Range');
        range.setProperty('Text', text);
    }

    /**
     * Add row
     * @returns {COMObject} New row
     */
    addRow() {
        const rows = this.getRows();
        return rows.invoke('Add');
    }

    /**
     * Add column
     * @returns {COMObject} New column
     */
    addColumn() {
        const columns = this.getColumns();
        return columns.invoke('Add');
    }

    /**
     * Delete row
     * @param {number} index - Row index (1-based)
     */
    deleteRow(index) {
        const rows = this.getRows();
        const row = rows.invoke('Item', index);
        row.invoke('Delete');
    }

    /**
     * Delete column
     * @param {number} index - Column index (1-based)
     */
    deleteColumn(index) {
        const columns = this.getColumns();
        const column = columns.invoke('Item', index);
        column.invoke('Delete');
    }

    /**
     * Delete table
     */
    delete() {
        this.table.invoke('Delete');
    }

    /**
     * Auto format table
     * @param {number} format - Format constant
     */
    autoFormat(format) {
        this.table.invoke('AutoFormat', format);
    }

    /**
     * Release COM object
     */
    release() {
        this.table.release();
    }
}

/**
 * @class WordFont
 * @description Wrapper for Word Font
 */
class WordFont {
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
     * @returns {number} Underline style
     */
    getUnderline() {
        return this.font.getProperty('Underline');
    }

    /**
     * Set underline
     * @param {number} value - Underline style constant
     */
    setUnderline(value) {
        this.font.setProperty('Underline', value);
    }

    /**
     * Get color
     * @returns {number} Color value
     */
    getColor() {
        return this.font.getProperty('Color');
    }

    /**
     * Set color
     * @param {number} value - Color value
     */
    setColor(value) {
        this.font.setProperty('Color', value);
    }

    /**
     * Get subscript
     * @returns {boolean} True if subscript
     */
    getSubscript() {
        return this.font.getProperty('Subscript');
    }

    /**
     * Set subscript
     * @param {boolean} value - True for subscript
     */
    setSubscript(value) {
        this.font.setProperty('Subscript', value ? -1 : 0);
    }

    /**
     * Get superscript
     * @returns {boolean} True if superscript
     */
    getSuperscript() {
        return this.font.getProperty('Superscript');
    }

    /**
     * Set superscript
     * @param {boolean} value - True for superscript
     */
    setSuperscript(value) {
        this.font.setProperty('Superscript', value ? -1 : 0);
    }

    /**
     * Release COM object
     */
    release() {
        this.font.release();
    }
}

/**
 * @class WordPageSetup
 * @description Wrapper for Word PageSetup
 */
class WordPageSetup {
    constructor(comObject) {
        this.pageSetup = comObject;
    }

    /**
     * Get top margin
     * @returns {number} Top margin in points
     */
    getTopMargin() {
        return this.pageSetup.getProperty('TopMargin');
    }

    /**
     * Set top margin
     * @param {number} value - Top margin in points
     */
    setTopMargin(value) {
        this.pageSetup.setProperty('TopMargin', value);
    }

    /**
     * Get bottom margin
     * @returns {number} Bottom margin in points
     */
    getBottomMargin() {
        return this.pageSetup.getProperty('BottomMargin');
    }

    /**
     * Set bottom margin
     * @param {number} value - Bottom margin in points
     */
    setBottomMargin(value) {
        this.pageSetup.setProperty('BottomMargin', value);
    }

    /**
     * Get left margin
     * @returns {number} Left margin in points
     */
    getLeftMargin() {
        return this.pageSetup.getProperty('LeftMargin');
    }

    /**
     * Set left margin
     * @param {number} value - Left margin in points
     */
    setLeftMargin(value) {
        this.pageSetup.setProperty('LeftMargin', value);
    }

    /**
     * Get right margin
     * @returns {number} Right margin in points
     */
    getRightMargin() {
        return this.pageSetup.getProperty('RightMargin');
    }

    /**
     * Set right margin
     * @param {number} value - Right margin in points
     */
    setRightMargin(value) {
        this.pageSetup.setProperty('RightMargin', value);
    }

    /**
     * Get orientation
     * @returns {number} Orientation constant
     */
    getOrientation() {
        return this.pageSetup.getProperty('Orientation');
    }

    /**
     * Set orientation
     * @param {number} value - Orientation constant
     */
    setOrientation(value) {
        this.pageSetup.setProperty('Orientation', value);
    }

    /**
     * Get page width
     * @returns {number} Page width in points
     */
    getPageWidth() {
        return this.pageSetup.getProperty('PageWidth');
    }

    /**
     * Set page width
     * @param {number} value - Page width in points
     */
    setPageWidth(value) {
        this.pageSetup.setProperty('PageWidth', value);
    }

    /**
     * Get page height
     * @returns {number} Page height in points
     */
    getPageHeight() {
        return this.pageSetup.getProperty('PageHeight');
    }

    /**
     * Set page height
     * @param {number} value - Page height in points
     */
    setPageHeight(value) {
        this.pageSetup.setProperty('PageHeight', value);
    }

    /**
     * Release COM object
     */
    release() {
        this.pageSetup.release();
    }
}

module.exports = {
    WordConnector,
    WordDocument,
    WordRange,
    WordParagraph,
    WordParagraphFormat,
    WordTable,
    WordFont,
    WordPageSetup,
    WdSaveFormat,
    WdParagraphAlignment,
    WdOrientation,
    WdUnderline,
    WdColor,
    WdBreakType,
};
