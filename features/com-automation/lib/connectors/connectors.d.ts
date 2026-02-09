/**
 * @fileoverview TypeScript definitions for COM Automation Connectors
 * Provides IntelliSense support for all Office and SAP automation features
 */

// ============================================================================
// EXCEL CONNECTOR
// ============================================================================

/**
 * Excel file format constants
 */
export enum XlFileFormat {
    xlWorkbookDefault = 51,
    xlOpenXMLWorkbook = 51,
    xlOpenXMLWorkbookMacroEnabled = 52,
    xlExcel8 = 56,
    xlCSV = 6,
    xlTextWindows = 20,
    xlHtml = 44
}

/**
 * Excel chart type constants
 */
export enum XlChartType {
    xlColumnClustered = 51,
    xlColumnStacked = 52,
    xlLine = 4,
    xlLineMarkers = 65,
    xlPie = 5,
    xlBarClustered = 57,
    xlArea = 1,
    xlXYScatter = -4169
}

/**
 * Excel calculation mode constants
 */
export enum XlCalculation {
    xlCalculationAutomatic = -4105,
    xlCalculationManual = -4135,
    xlCalculationSemiautomatic = 2
}

/**
 * Excel horizontal alignment constants
 */
export enum XlHAlign {
    xlHAlignCenter = -4108,
    xlHAlignLeft = -4131,
    xlHAlignRight = -4152,
    xlHAlignJustify = -4130,
    xlHAlignDistributed = -4117
}

/**
 * Excel vertical alignment constants
 */
export enum XlVAlign {
    xlVAlignCenter = -4108,
    xlVAlignTop = -4160,
    xlVAlignBottom = -4107,
    xlVAlignJustify = -4130,
    xlVAlignDistributed = -4117
}

/**
 * Excel Connector - High-level API for Microsoft Excel automation
 * 
 * @example
 * ```typescript
 * const { ExcelConnector } = require('node-winautomation');
 * const excel = new ExcelConnector();
 * 
 * // Open workbook
 * const workbook = excel.openWorkbook('C:\\data.xlsx');
 * const sheet = workbook.getWorksheet(1);
 * 
 * // Write data
 * sheet.writeCell(1, 1, 'Hello');
 * sheet.writeRange('A2:B3', [[1, 2], [3, 4]]);
 * 
 * // Create pivot table
 * sheet.addPivotTable('A1:D10', 'F1', 'SalesPivot');
 * 
 * // Save and close
 * workbook.save();
 * workbook.close();
 * excel.quit();
 * ```
 */
export class ExcelConnector {
    /**
     * Creates a new Excel application instance
     * @throws {Error} If Excel is not installed or COM initialization fails
     */
    constructor();

    /**
     * Sets the visibility of the Excel application window
     * @param visible - True to show Excel window, false to hide
     * @example
     * ```typescript
     * excel.setVisible(true); // Show Excel window
     * ```
     */
    setVisible(visible: boolean): void;

    /**
     * Gets the visibility state of the Excel application
     * @returns True if Excel window is visible
     */
    getVisible(): boolean;

    /**
     * Creates a new blank workbook
     * @returns Workbook wrapper object
     * @example
     * ```typescript
     * const workbook = excel.addWorkbook();
     * const sheet = workbook.getWorksheet(1);
     * ```
     */
    addWorkbook(): Workbook;

    /**
     * Opens an existing workbook from file
     * @param path - Full path to the Excel file
     * @param readOnly - Open in read-only mode (default: false)
     * @returns Workbook wrapper object
     * @throws {Error} If file doesn't exist or cannot be opened
     * @example
     * ```typescript
     * const workbook = excel.openWorkbook('C:\\data.xlsx');
     * const readOnlyWb = excel.openWorkbook('C:\\data.xlsx', true);
     * ```
     */
    openWorkbook(path: string, readOnly?: boolean): Workbook;

    /**
     * Gets the currently active workbook
     * @returns Active workbook wrapper or null if no workbook is active
     */
    getActiveWorkbook(): Workbook | null;

    /**
     * Gets the number of open workbooks
     * @returns Count of open workbooks
     */
    getWorkbookCount(): number;

    /**
     * Sets the calculation mode for Excel
     * @param mode - Calculation mode constant
     * @example
     * ```typescript
     * excel.setCalculation(XlCalculation.xlCalculationManual);
     * ```
     */
    setCalculation(mode: XlCalculation): void;

    /**
     * Gets the current calculation mode
     * @returns Current calculation mode constant
     */
    getCalculation(): XlCalculation;

    /**
     * Gets the Excel application version
     * @returns Version string (e.g., "16.0" for Excel 2016)
     */
    getVersion(): string;

    /**
     * Quits the Excel application
     * @param saveChanges - Save all open workbooks before quitting (default: false)
     * @example
     * ```typescript
     * excel.quit(true); // Save all and quit
     * ```
     */
    quit(saveChanges?: boolean): void;

    /**
     * Releases COM resources
     * @remarks Always call this when done to prevent memory leaks
     */
    release(): void;
}

/**
 * Excel Workbook wrapper
 */
export class Workbook {
    /**
     * Gets the workbook name
     * @returns Workbook filename
     */
    getName(): string;

    /**
     * Gets the full path of the workbook
     * @returns Full file path or empty string if not saved
     */
    getFullName(): string;

    /**
     * Gets the number of worksheets in the workbook
     * @returns Worksheet count
     */
    getWorksheetCount(): number;

    /**
     * Gets a worksheet by index (1-based)
     * @param index - Worksheet index (1-based)
     * @returns Worksheet wrapper object
     * @throws {Error} If index is out of range
     * @example
     * ```typescript
     * const sheet = workbook.getWorksheet(1); // First sheet
     * ```
     */
    getWorksheet(index: number): Worksheet;

    /**
     * Gets a worksheet by name
     * @param name - Worksheet name
     * @returns Worksheet wrapper object
     * @throws {Error} If worksheet doesn't exist
     * @example
     * ```typescript
     * const sheet = workbook.getWorksheetByName('Sales');
     * ```
     */
    getWorksheetByName(name: string): Worksheet;

    /**
     * Adds a new worksheet
     * @param name - Optional worksheet name
     * @returns New worksheet wrapper object
     * @example
     * ```typescript
     * const sheet = workbook.addWorksheet('NewSheet');
     * ```
     */
    addWorksheet(name?: string): Worksheet;

    /**
     * Saves the workbook
     * @throws {Error} If workbook hasn't been saved before (use saveAs instead)
     */
    save(): void;

    /**
     * Saves the workbook with a new name or format
     * @param path - Full path for the file
     * @param format - File format constant (default: xlWorkbookDefault)
     * @example
     * ```typescript
     * workbook.saveAs('C:\\output.xlsx');
     * workbook.saveAs('C:\\output.csv', XlFileFormat.xlCSV);
     * ```
     */
    saveAs(path: string, format?: XlFileFormat): void;

    /**
     * Closes the workbook
     * @param saveChanges - Save before closing (default: false)
     */
    close(saveChanges?: boolean): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Excel Worksheet wrapper
 */
export class Worksheet {
    /**
     * Gets the worksheet name
     * @returns Worksheet name
     */
    getName(): string;

    /**
     * Sets the worksheet name
     * @param name - New worksheet name
     * @throws {Error} If name is invalid or already exists
     */
    setName(name: string): void;

    /**
     * Writes a value to a single cell
     * @param row - Row number (1-based)
     * @param col - Column number (1-based)
     * @param value - Value to write (string, number, or formula)
     * @example
     * ```typescript
     * sheet.writeCell(1, 1, 'Hello');
     * sheet.writeCell(2, 1, 42);
     * sheet.writeCell(3, 1, '=SUM(A1:A2)');
     * ```
     */
    writeCell(row: number, col: number, value: string | number): void;

    /**
     * Reads a value from a single cell
     * @param row - Row number (1-based)
     * @param col - Column number (1-based)
     * @returns Cell value
     * @example
     * ```typescript
     * const value = sheet.readCell(1, 1);
     * ```
     */
    readCell(row: number, col: number): any;

    /**
     * Writes a 2D array to a range
     * @param range - Range address (e.g., "A1:C3")
     * @param data - 2D array of values
     * @example
     * ```typescript
     * sheet.writeRange('A1:B2', [
     *   ['Name', 'Age'],
     *   ['John', 30]
     * ]);
     * ```
     */
    writeRange(range: string, data: any[][]): void;

    /**
     * Reads values from a range
     * @param range - Range address (e.g., "A1:C3")
     * @returns 2D array of values
     * @example
     * ```typescript
     * const data = sheet.readRange('A1:C10');
     * ```
     */
    readRange(range: string): any[][];

    /**
     * Gets a Range object for manipulation
     * @param address - Range address (e.g., "A1:C3")
     * @returns Range wrapper object
     * @example
     * ```typescript
     * const range = sheet.getRange('A1:C3');
     * range.setBold(true);
     * range.setBackgroundColor(255, 255, 0);
     * ```
     */
    getRange(address: string): Range;

    /**
     * Gets the last used row number
     * @returns Last row with data
     */
    getLastRow(): number;

    /**
     * Gets the last used column number
     * @returns Last column with data
     */
    getLastColumn(): number;

    /**
     * Appends a row of data at the end
     * @param data - Array of values
     * @example
     * ```typescript
     * sheet.appendRow(['John', 30, 'Sales']);
     * ```
     */
    appendRow(data: any[]): void;

    /**
     * Creates a pivot table
     * @param sourceRange - Source data range (e.g., "A1:D10")
     * @param destinationCell - Where to place pivot table (e.g., "F1")
     * @param tableName - Name for the pivot table
     * @returns PivotTable wrapper object
     * @example
     * ```typescript
     * const pivot = sheet.addPivotTable('A1:D100', 'F1', 'SalesPivot');
     * pivot.addRowField('Product');
     * pivot.addDataField('Sales', -4157); // xlSum
     * pivot.refresh();
     * ```
     */
    addPivotTable(sourceRange: string, destinationCell: string, tableName: string): PivotTable;

    /**
     * Gets an existing pivot table by name
     * @param name - Pivot table name
     * @returns PivotTable wrapper object
     */
    getPivotTable(name: string): PivotTable;

    /**
     * Gets the number of pivot tables in the worksheet
     * @returns Pivot table count
     */
    getPivotTableCount(): number;

    /**
     * Deletes the worksheet
     * @remarks Cannot be undone
     */
    delete(): void;

    /**
     * Copies the worksheet
     * @param afterSheet - Optional worksheet to copy after
     * @returns New worksheet wrapper
     */
    copy(afterSheet?: Worksheet): Worksheet;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Excel Range wrapper for formatting and manipulation
 */
export class Range {
    /**
     * Sets bold formatting
     * @param bold - True for bold
     */
    setBold(bold: boolean): void;

    /**
     * Sets italic formatting
     * @param italic - True for italic
     */
    setItalic(italic: boolean): void;

    /**
     * Sets font size
     * @param size - Font size in points
     */
    setFontSize(size: number): void;

    /**
     * Sets font name
     * @param name - Font name (e.g., "Arial")
     */
    setFontName(name: string): void;

    /**
     * Sets font color
     * @param r - Red component (0-255)
     * @param g - Green component (0-255)
     * @param b - Blue component (0-255)
     */
    setFontColor(r: number, g: number, b: number): void;

    /**
     * Sets background color
     * @param r - Red component (0-255)
     * @param g - Green component (0-255)
     * @param b - Blue component (0-255)
     */
    setBackgroundColor(r: number, g: number, b: number): void;

    /**
     * Sets horizontal alignment
     * @param alignment - Alignment constant
     */
    setHorizontalAlignment(alignment: XlHAlign): void;

    /**
     * Sets vertical alignment
     * @param alignment - Alignment constant
     */
    setVerticalAlignment(alignment: XlVAlign): void;

    /**
     * Adds data validation to the range
     * @param type - Validation type (1=whole number, 3=list, etc.)
     * @param operator - Operator (1=between, 3=equal, etc.)
     * @param formula1 - First formula/value
     * @param formula2 - Second formula/value (optional)
     * @example
     * ```typescript
     * // Dropdown list
     * range.addValidation(3, 1, 'Yes,No,Maybe');
     * 
     * // Number between 1 and 100
     * range.addValidation(1, 1, '1', '100');
     * ```
     */
    addValidation(type: number, operator: number, formula1: string, formula2?: string): void;

    /**
     * Clears data validation from the range
     */
    clearValidation(): void;

    /**
     * Adds conditional formatting to the range
     * @param type - Format condition type
     * @param operator - Comparison operator
     * @param formula1 - First formula/value
     * @param formula2 - Second formula/value (optional)
     * @returns Format condition object
     * @example
     * ```typescript
     * // Highlight cells > 100
     * const condition = range.addConditionalFormatting(1, 5, '100');
     * ```
     */
    addConditionalFormatting(type: number, operator: number, formula1: string, formula2?: string): any;

    /**
     * Clears all conditional formatting from the range
     */
    clearConditionalFormatting(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Excel Pivot Table wrapper
 */
export class PivotTable {
    /**
     * Gets the pivot table name
     * @returns Pivot table name
     */
    getName(): string;

    /**
     * Sets the pivot table name
     * @param name - New name
     */
    setName(name: string): void;

    /**
     * Adds a field to the row area
     * @param fieldName - Field name from source data
     * @param position - Optional position (1-based)
     * @example
     * ```typescript
     * pivot.addRowField('Product');
     * pivot.addRowField('Category', 1); // Add at position 1
     * ```
     */
    addRowField(fieldName: string, position?: number): void;

    /**
     * Adds a field to the column area
     * @param fieldName - Field name from source data
     * @param position - Optional position (1-based)
     */
    addColumnField(fieldName: string, position?: number): void;

    /**
     * Adds a field to the data area
     * @param fieldName - Field name from source data
     * @param functionType - Aggregation function (-4157=Sum, -4106=Count, -4106=Average)
     * @param caption - Optional custom caption
     * @example
     * ```typescript
     * pivot.addDataField('Sales', -4157, 'Total Sales'); // Sum
     * pivot.addDataField('Quantity', -4106, 'Count'); // Count
     * ```
     */
    addDataField(fieldName: string, functionType?: number, caption?: string): void;

    /**
     * Adds a field to the page/filter area
     * @param fieldName - Field name from source data
     * @param position - Optional position (1-based)
     */
    addPageField(fieldName: string, position?: number): void;

    /**
     * Refreshes the pivot table data
     */
    refresh(): void;

    /**
     * Clears the pivot table
     */
    clear(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

// ============================================================================
// WORD CONNECTOR
// ============================================================================

/**
 * Word save format constants
 */
export enum WdSaveFormat {
    wdFormatDocument = 0,
    wdFormatDocumentDefault = 16,
    wdFormatPDF = 17,
    wdFormatRTF = 6,
    wdFormatText = 2,
    wdFormatHTML = 8
}

/**
 * Word paragraph alignment constants
 */
export enum WdParagraphAlignment {
    wdAlignParagraphLeft = 0,
    wdAlignParagraphCenter = 1,
    wdAlignParagraphRight = 2,
    wdAlignParagraphJustify = 3
}

/**
 * Word Connector - High-level API for Microsoft Word automation
 * 
 * @remarks
 * Note: Table, bookmark, and comment operations have known COM limitations (73% success rate).
 * See WORD_COM_LIMITATIONS.md for details and workarounds.
 * 
 * @example
 * ```typescript
 * const { WordConnector } = require('node-winautomation');
 * const word = new WordConnector();
 * 
 * // Create document
 * const doc = word.addDocument();
 * doc.setText('Hello World');
 * doc.addParagraph('New paragraph');
 * 
 * // Format text
 * const range = doc.getContent();
 * range.setBold(true);
 * range.setFontSize(14);
 * 
 * // Save and close
 * doc.saveAs('C:\\document.docx');
 * doc.close();
 * word.quit();
 * ```
 */
export class WordConnector {
    /**
     * Creates a new Word application instance
     * @throws {Error} If Word is not installed or COM initialization fails
     */
    constructor();

    /**
     * Sets the visibility of the Word application window
     * @param visible - True to show Word window
     */
    setVisible(visible: boolean): void;

    /**
     * Gets the visibility state of the Word application
     * @returns True if Word window is visible
     */
    getVisible(): boolean;

    /**
     * Creates a new blank document
     * @returns Document wrapper object
     */
    addDocument(): WordDocument;

    /**
     * Opens an existing document from file
     * @param path - Full path to the Word file
     * @param readOnly - Open in read-only mode (default: false)
     * @returns Document wrapper object
     * @throws {Error} If file doesn't exist or cannot be opened
     */
    openDocument(path: string, readOnly?: boolean): WordDocument;

    /**
     * Gets the currently active document
     * @returns Active document wrapper or null
     */
    getActiveDocument(): WordDocument | null;

    /**
     * Gets the number of open documents
     * @returns Count of open documents
     */
    getDocumentCount(): number;

    /**
     * Gets the Word application version
     * @returns Version string
     */
    getVersion(): string;

    /**
     * Quits the Word application
     * @param saveChanges - Save all open documents before quitting
     */
    quit(saveChanges?: boolean): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Word Document wrapper
 */
export class WordDocument {
    /**
     * Gets the document name
     * @returns Document filename
     */
    getName(): string;

    /**
     * Gets the full path of the document
     * @returns Full file path or empty string if not saved
     */
    getFullName(): string;

    /**
     * Sets the entire document text
     * @param text - Text content
     * @remarks This replaces all existing content
     */
    setText(text: string): void;

    /**
     * Gets the entire document text
     * @returns Document text content
     */
    getText(): string;

    /**
     * Gets the document content as a Range object
     * @returns Range wrapper for the entire document
     */
    getContent(): WordRange;

    /**
     * Adds a new paragraph at the end
     * @param text - Paragraph text
     * @returns Paragraph wrapper object
     */
    addParagraph(text: string): WordParagraph;

    /**
     * Gets a paragraph by index
     * @param index - Paragraph index (1-based)
     * @returns Paragraph wrapper object
     */
    getParagraph(index: number): WordParagraph;

    /**
     * Gets the number of paragraphs
     * @returns Paragraph count
     */
    getParagraphCount(): number;

    /**
     * Adds a table to the document
     * @param rows - Number of rows
     * @param cols - Number of columns
     * @param range - Optional range where table should be inserted
     * @returns Table wrapper object
     * @experimental This method has known COM limitations (73% success rate)
     * @see WORD_COM_LIMITATIONS.md for details and workarounds
     */
    addTable(rows: number, cols: number, range?: WordRange): WordTable;

    /**
     * Gets a table by index
     * @param index - Table index (1-based)
     * @returns Table wrapper object
     */
    getTable(index: number): WordTable;

    /**
     * Gets the number of tables
     * @returns Table count
     */
    getTableCount(): number;

    /**
     * Adds a bookmark
     * @param name - Bookmark name
     * @param range - Optional range to bookmark
     * @experimental This method has known COM limitations (73% success rate)
     * @see WORD_COM_LIMITATIONS.md for details and workarounds
     */
    addBookmark(name: string, range?: WordRange): void;

    /**
     * Checks if a bookmark exists
     * @param name - Bookmark name
     * @returns True if bookmark exists
     * @experimental This method has known COM limitations
     */
    bookmarkExists(name: string): boolean;

    /**
     * Navigates to a bookmark
     * @param name - Bookmark name
     * @returns Range at bookmark location
     * @experimental This method has known COM limitations
     */
    goToBookmark(name: string): WordRange;

    /**
     * Adds a comment
     * @param text - Comment text
     * @param range - Optional range to comment on
     * @experimental This method has known COM limitations (73% success rate)
     * @see WORD_COM_LIMITATIONS.md for details and workarounds
     */
    addComment(text: string, range?: WordRange): void;

    /**
     * Finds text in the document
     * @param text - Text to find
     * @returns True if found
     */
    find(text: string): boolean;

    /**
     * Replaces text in the document
     * @param findText - Text to find
     * @param replaceText - Replacement text
     * @param replaceAll - Replace all occurrences (default: true)
     * @returns Number of replacements made
     */
    replace(findText: string, replaceText: string, replaceAll?: boolean): number;

    /**
     * Saves the document
     * @throws {Error} If document hasn't been saved before (use saveAs instead)
     */
    save(): void;

    /**
     * Saves the document with a new name or format
     * @param path - Full path for the file
     * @param format - File format constant (default: wdFormatDocumentDefault)
     */
    saveAs(path: string, format?: WdSaveFormat): void;

    /**
     * Exports the document to PDF
     * @param path - Full path for the PDF file
     */
    exportToPDF(path: string): void;

    /**
     * Closes the document
     * @param saveChanges - Save before closing
     */
    close(saveChanges?: boolean): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Word Range wrapper
 */
export class WordRange {
    /**
     * Gets the text content
     * @returns Text content
     */
    getText(): string;

    /**
     * Sets the text content
     * @param text - Text to set
     */
    setText(text: string): void;

    /**
     * Sets bold formatting
     * @param bold - True for bold
     */
    setBold(bold: boolean): void;

    /**
     * Sets italic formatting
     * @param italic - True for italic
     */
    setItalic(italic: boolean): void;

    /**
     * Sets font size
     * @param size - Font size in points
     */
    setFontSize(size: number): void;

    /**
     * Sets font name
     * @param name - Font name
     */
    setFontName(name: string): void;

    /**
     * Sets font color
     * @param r - Red component (0-255)
     * @param g - Green component (0-255)
     * @param b - Blue component (0-255)
     */
    setFontColor(r: number, g: number, b: number): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Word Paragraph wrapper
 */
export class WordParagraph {
    /**
     * Gets the paragraph alignment
     * @returns Alignment constant
     */
    getAlignment(): WdParagraphAlignment;

    /**
     * Sets the paragraph alignment
     * @param alignment - Alignment constant
     */
    setAlignment(alignment: WdParagraphAlignment): void;

    /**
     * Gets the paragraph range
     * @returns Range wrapper
     */
    getRange(): WordRange;

    /**
     * Deletes the paragraph
     */
    delete(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Word Table wrapper
 */
export class WordTable {
    /**
     * Gets the number of rows
     * @returns Row count
     */
    getRowCount(): number;

    /**
     * Gets the number of columns
     * @returns Column count
     */
    getColumnCount(): number;

    /**
     * Gets cell text
     * @param row - Row index (1-based)
     * @param col - Column index (1-based)
     * @returns Cell text
     */
    getCellText(row: number, col: number): string;

    /**
     * Sets cell text
     * @param row - Row index (1-based)
     * @param col - Column index (1-based)
     * @param text - Text to set
     */
    setCellText(row: number, col: number, text: string): void;

    /**
     * Adds a row at the end
     * @returns New row object
     */
    addRow(): any;

    /**
     * Adds a column at the end
     * @returns New column object
     */
    addColumn(): any;

    /**
     * Deletes the table
     */
    delete(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

// ============================================================================
// OUTLOOK CONNECTOR
// ============================================================================

/**
 * Outlook item type constants
 */
export enum OlItemType {
    olMailItem = 0,
    olAppointmentItem = 1,
    olContactItem = 2,
    olTaskItem = 3,
    olJournalItem = 4,
    olNoteItem = 5,
    olPostItem = 6
}

/**
 * Outlook default folder constants
 */
export enum OlDefaultFolders {
    olFolderInbox = 6,
    olFolderSentMail = 5,
    olFolderDrafts = 16,
    olFolderDeletedItems = 3,
    olFolderOutbox = 4,
    olFolderCalendar = 9,
    olFolderContacts = 10,
    olFolderTasks = 13
}

/**
 * Outlook Connector - High-level API for Microsoft Outlook automation
 * 
 * @example
 * ```typescript
 * const { OutlookConnector } = require('node-winautomation');
 * const outlook = new OutlookConnector();
 * 
 * // Send email
 * const mail = outlook.createMailItem();
 * mail.setTo('user@example.com');
 * mail.setSubject('Test Email');
 * mail.setBody('Hello from Node.js!');
 * mail.send();
 * 
 * // Create appointment
 * const appt = outlook.createAppointmentItem();
 * appt.setSubject('Meeting');
 * appt.setStart(new Date());
 * appt.setDuration(60);
 * appt.save();
 * 
 * // Manage categories
 * mail.addCategory('Important');
 * mail.setVotingOptions('Yes;No;Maybe');
 * ```
 */
export class OutlookConnector {
    /**
     * Creates a new Outlook application instance
     * @throws {Error} If Outlook is not installed or COM initialization fails
     */
    constructor();

    /**
     * Creates a new mail item
     * @returns MailItem wrapper object
     */
    createMailItem(): MailItem;

    /**
     * Creates a new appointment item
     * @returns AppointmentItem wrapper object
     */
    createAppointmentItem(): AppointmentItem;

    /**
     * Creates a new contact item
     * @returns ContactItem wrapper object
     */
    createContactItem(): ContactItem;

    /**
     * Creates a new task item
     * @returns TaskItem wrapper object
     */
    createTaskItem(): TaskItem;

    /**
     * Gets a default folder
     * @param folderType - Folder type constant
     * @returns Folder wrapper object
     */
    getDefaultFolder(folderType: OlDefaultFolders): Folder;

    /**
     * Gets the Inbox folder
     * @returns Folder wrapper object
     */
    getInbox(): Folder;

    /**
     * Gets the Sent Mail folder
     * @returns Folder wrapper object
     */
    getSentMail(): Folder;

    /**
     * Gets the Drafts folder
     * @returns Folder wrapper object
     */
    getDrafts(): Folder;

    /**
     * Gets Outlook rules collection
     * @returns Rules collection object
     */
    getRules(): any;

    /**
     * Creates a new rule
     * @param name - Rule name
     * @param ruleType - Rule type (default: 0)
     * @returns OutlookRule wrapper object
     */
    createRule(name: string, ruleType?: number): OutlookRule;

    /**
     * Gets an existing rule by name
     * @param name - Rule name
     * @returns OutlookRule wrapper object
     */
    getRule(name: string): OutlookRule;

    /**
     * Saves all rules
     */
    saveRules(): void;

    /**
     * Quits the Outlook application
     */
    quit(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Outlook Mail Item wrapper
 */
export class MailItem {
    /**
     * Gets the subject
     * @returns Email subject
     */
    getSubject(): string;

    /**
     * Sets the subject
     * @param subject - Email subject
     */
    setSubject(subject: string): void;

    /**
     * Gets the body text
     * @returns Email body
     */
    getBody(): string;

    /**
     * Sets the body text
     * @param body - Email body
     */
    setBody(body: string): void;

    /**
     * Gets the HTML body
     * @returns HTML body
     */
    getHTMLBody(): string;

    /**
     * Sets the HTML body
     * @param html - HTML body
     */
    setHTMLBody(html: string): void;

    /**
     * Gets the To recipients
     * @returns Semicolon-separated email addresses
     */
    getTo(): string;

    /**
     * Sets the To recipients
     * @param to - Semicolon-separated email addresses
     */
    setTo(to: string): void;

    /**
     * Gets the CC recipients
     * @returns Semicolon-separated email addresses
     */
    getCC(): string;

    /**
     * Sets the CC recipients
     * @param cc - Semicolon-separated email addresses
     */
    setCC(cc: string): void;

    /**
     * Gets the BCC recipients
     * @returns Semicolon-separated email addresses
     */
    getBCC(): string;

    /**
     * Sets the BCC recipients
     * @param bcc - Semicolon-separated email addresses
     */
    setBCC(bcc: string): void;

    /**
     * Gets the categories
     * @returns Semicolon-separated categories
     */
    getCategories(): string;

    /**
     * Sets the categories
     * @param categories - Semicolon-separated categories
     */
    setCategories(categories: string): void;

    /**
     * Adds a category to the mail item
     * @param category - Category name
     */
    addCategory(category: string): void;

    /**
     * Removes a category from the mail item
     * @param category - Category name
     */
    removeCategory(category: string): void;

    /**
     * Clears all categories
     */
    clearCategories(): void;

    /**
     * Sets voting options
     * @param options - Semicolon-separated options (e.g., "Yes;No;Maybe")
     */
    setVotingOptions(options: string): void;

    /**
     * Gets voting options
     * @returns Semicolon-separated options
     */
    getVotingOptions(): string;

    /**
     * Gets the voting response
     * @returns Selected voting option
     */
    getVotingResponse(): string;

    /**
     * Sends the email
     */
    send(): void;

    /**
     * Saves the email
     */
    save(): void;

    /**
     * Displays the email in a window
     */
    display(): void;

    /**
     * Closes the email
     * @param saveChanges - Save before closing
     */
    close(saveChanges?: boolean): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Outlook Appointment Item wrapper
 */
export class AppointmentItem {
    /**
     * Gets the subject
     * @returns Appointment subject
     */
    getSubject(): string;

    /**
     * Sets the subject
     * @param subject - Appointment subject
     */
    setSubject(subject: string): void;

    /**
     * Gets the start time
     * @returns Start date/time
     */
    getStart(): Date;

    /**
     * Sets the start time
     * @param start - Start date/time
     */
    setStart(start: Date): void;

    /**
     * Gets the end time
     * @returns End date/time
     */
    getEnd(): Date;

    /**
     * Sets the end time
     * @param end - End date/time
     */
    setEnd(end: Date): void;

    /**
     * Gets the duration in minutes
     * @returns Duration in minutes
     */
    getDuration(): number;

    /**
     * Sets the duration in minutes
     * @param minutes - Duration in minutes
     */
    setDuration(minutes: number): void;

    /**
     * Gets the location
     * @returns Location string
     */
    getLocation(): string;

    /**
     * Sets the location
     * @param location - Location string
     */
    setLocation(location: string): void;

    /**
     * Saves the appointment
     */
    save(): void;

    /**
     * Displays the appointment in a window
     */
    display(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Outlook Contact Item wrapper
 */
export class ContactItem {
    /**
     * Gets the full name
     * @returns Full name
     */
    getFullName(): string;

    /**
     * Sets the full name
     * @param name - Full name
     */
    setFullName(name: string): void;

    /**
     * Gets the email address
     * @returns Email address
     */
    getEmail1Address(): string;

    /**
     * Sets the email address
     * @param email - Email address
     */
    setEmail1Address(email: string): void;

    /**
     * Gets the business phone
     * @returns Phone number
     */
    getBusinessTelephoneNumber(): string;

    /**
     * Sets the business phone
     * @param phone - Phone number
     */
    setBusinessTelephoneNumber(phone: string): void;

    /**
     * Saves the contact
     */
    save(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Outlook Task Item wrapper
 */
export class TaskItem {
    /**
     * Gets the subject
     * @returns Task subject
     */
    getSubject(): string;

    /**
     * Sets the subject
     * @param subject - Task subject
     */
    setSubject(subject: string): void;

    /**
     * Gets the due date
     * @returns Due date
     */
    getDueDate(): Date;

    /**
     * Sets the due date
     * @param date - Due date
     */
    setDueDate(date: Date): void;

    /**
     * Marks the task as complete
     */
    markComplete(): void;

    /**
     * Saves the task
     */
    save(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Outlook Folder wrapper
 */
export class Folder {
    /**
     * Gets the folder name
     * @returns Folder name
     */
    getName(): string;

    /**
     * Gets the number of items in the folder
     * @returns Item count
     */
    getCount(): number;

    /**
     * Gets the number of unread items
     * @returns Unread item count
     */
    getUnReadItemCount(): number;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * Outlook Rule wrapper
 */
export class OutlookRule {
    /**
     * Gets the rule name
     * @returns Rule name
     */
    getName(): string;

    /**
     * Sets the rule name
     * @param name - Rule name
     */
    setName(name: string): void;

    /**
     * Gets whether the rule is enabled
     * @returns True if enabled
     */
    getEnabled(): boolean;

    /**
     * Sets whether the rule is enabled
     * @param enabled - True to enable
     */
    setEnabled(enabled: boolean): void;

    /**
     * Gets the rule conditions
     * @returns Conditions object
     */
    getConditions(): any;

    /**
     * Gets the rule actions
     * @returns Actions object
     */
    getActions(): any;

    /**
     * Executes the rule
     * @param showProgress - Show progress dialog
     * @param folder - Optional folder to execute on
     */
    execute(showProgress?: boolean, folder?: Folder): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

// ============================================================================
// POWERPOINT CONNECTOR
// ============================================================================

/**
 * PowerPoint save format constants
 */
export enum PpSaveAsFileType {
    ppSaveAsPresentation = 1,
    ppSaveAsPDF = 32,
    ppSaveAsOpenXMLPresentation = 24,
    ppSaveAsShow = 7
}

/**
 * PowerPoint Connector - High-level API for Microsoft PowerPoint automation
 * 
 * @example
 * ```typescript
 * const { PowerPointConnector } = require('node-winautomation');
 * const ppt = new PowerPointConnector();
 * 
 * // Create presentation
 * const pres = ppt.addPresentation();
 * const slide = pres.addSlide(1); // Blank layout
 * 
 * // Add content
 * const textbox = slide.addTextBox(100, 100, 400, 100);
 * textbox.setText('Hello PowerPoint!');
 * 
 * // Add transition
 * slide.setTransition(257, 2); // Fade, medium speed
 * 
 * // Save and close
 * pres.saveAs('C:\\presentation.pptx');
 * pres.close();
 * ppt.quit();
 * ```
 */
export class PowerPointConnector {
    /**
     * Creates a new PowerPoint application instance
     * @throws {Error} If PowerPoint is not installed or COM initialization fails
     */
    constructor();

    /**
     * Sets the visibility of the PowerPoint application window
     * @param visible - True to show PowerPoint window
     */
    setVisible(visible: boolean): void;

    /**
     * Gets the visibility state of the PowerPoint application
     * @returns True if PowerPoint window is visible
     */
    getVisible(): boolean;

    /**
     * Creates a new blank presentation
     * @returns Presentation wrapper object
     */
    addPresentation(): Presentation;

    /**
     * Opens an existing presentation from file
     * @param path - Full path to the PowerPoint file
     * @param readOnly - Open in read-only mode (default: false)
     * @returns Presentation wrapper object
     * @throws {Error} If file doesn't exist or cannot be opened
     */
    openPresentation(path: string, readOnly?: boolean): Presentation;

    /**
     * Gets the currently active presentation
     * @returns Active presentation wrapper or null
     */
    getActivePresentation(): Presentation | null;

    /**
     * Gets the number of open presentations
     * @returns Count of open presentations
     */
    getPresentationCount(): number;

    /**
     * Gets the PowerPoint application version
     * @returns Version string
     */
    getVersion(): string;

    /**
     * Quits the PowerPoint application
     */
    quit(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * PowerPoint Presentation wrapper
 */
export class Presentation {
    /**
     * Gets the presentation name
     * @returns Presentation filename
     */
    getName(): string;

    /**
     * Gets the full path of the presentation
     * @returns Full file path or empty string if not saved
     */
    getFullName(): string;

    /**
     * Gets the number of slides
     * @returns Slide count
     */
    getSlideCount(): number;

    /**
     * Gets a slide by index
     * @param index - Slide index (1-based)
     * @returns Slide wrapper object
     */
    getSlide(index: number): Slide;

    /**
     * Adds a new slide
     * @param layout - Slide layout (1=blank, 2=title, etc.)
     * @returns Slide wrapper object
     */
    addSlide(layout: number): Slide;

    /**
     * Saves the presentation
     * @throws {Error} If presentation hasn't been saved before (use saveAs instead)
     */
    save(): void;

    /**
     * Saves the presentation with a new name or format
     * @param path - Full path for the file
     * @param format - File format constant (default: ppSaveAsPresentation)
     */
    saveAs(path: string, format?: PpSaveAsFileType): void;

    /**
     * Closes the presentation
     */
    close(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * PowerPoint Slide wrapper
 */
export class Slide {
    /**
     * Gets the slide index
     * @returns Slide index (1-based)
     */
    getSlideIndex(): number;

    /**
     * Gets the number of shapes on the slide
     * @returns Shape count
     */
    getShapeCount(): number;

    /**
     * Gets a shape by index
     * @param index - Shape index (1-based)
     * @returns Shape wrapper object
     */
    getShape(index: number): Shape;

    /**
     * Adds a text box to the slide
     * @param left - Left position in points
     * @param top - Top position in points
     * @param width - Width in points
     * @param height - Height in points
     * @returns Shape wrapper object
     */
    addTextBox(left: number, top: number, width: number, height: number): Shape;

    /**
     * Adds a shape to the slide
     * @param shapeType - Shape type constant
     * @param left - Left position in points
     * @param top - Top position in points
     * @param width - Width in points
     * @param height - Height in points
     * @returns Shape wrapper object
     */
    addShape(shapeType: number, left: number, top: number, width: number, height: number): Shape;

    /**
     * Sets slide transition effect
     * @param type - Transition type constant
     * @param speed - Transition speed (1=fast, 2=medium, 3=slow)
     * @example
     * ```typescript
     * slide.setTransition(257, 2); // Fade, medium speed
     * ```
     */
    setTransition(type: number, speed?: number): void;

    /**
     * Sets transition duration
     * @param seconds - Duration in seconds
     */
    setTransitionDuration(seconds: number): void;

    /**
     * Sets whether slide advances automatically
     * @param advanceOnTime - True to advance automatically
     * @param advanceTime - Time in seconds before advancing
     */
    setTransitionAdvanceOnTime(advanceOnTime: boolean, advanceTime?: number): void;

    /**
     * Sets whether slide advances on click
     * @param advanceOnClick - True to advance on click
     */
    setTransitionAdvanceOnClick(advanceOnClick: boolean): void;

    /**
     * Gets slide notes
     * @returns Notes text
     * @experimental This method may have COM limitations in some PowerPoint versions
     */
    getNotes(): string;

    /**
     * Sets slide notes
     * @param text - Notes text
     * @experimental This method may have COM limitations in some PowerPoint versions
     */
    setNotes(text: string): void;

    /**
     * Deletes the slide
     */
    delete(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

/**
 * PowerPoint Shape wrapper
 */
export class Shape {
    /**
     * Gets the shape name
     * @returns Shape name
     */
    getName(): string;

    /**
     * Sets the shape name
     * @param name - Shape name
     */
    setName(name: string): void;

    /**
     * Gets the text content
     * @returns Text content
     */
    getText(): string;

    /**
     * Sets the text content
     * @param text - Text to set
     */
    setText(text: string): void;

    /**
     * Gets the left position
     * @returns Left position in points
     */
    getLeft(): number;

    /**
     * Sets the left position
     * @param left - Left position in points
     */
    setLeft(left: number): void;

    /**
     * Gets the top position
     * @returns Top position in points
     */
    getTop(): number;

    /**
     * Sets the top position
     * @param top - Top position in points
     */
    setTop(top: number): void;

    /**
     * Gets the width
     * @returns Width in points
     */
    getWidth(): number;

    /**
     * Sets the width
     * @param width - Width in points
     */
    setWidth(width: number): void;

    /**
     * Gets the height
     * @returns Height in points
     */
    getHeight(): number;

    /**
     * Sets the height
     * @param height - Height in points
     */
    setHeight(height: number): void;

    /**
     * Adds an animation effect to the shape
     * @param effect - Animation effect constant
     * @param trigger - Trigger type (0=on click, 1=with previous, 2=after previous)
     * @returns Animation effect object
     * @example
     * ```typescript
     * shape.addAnimation(10, 0); // Fade in on click
     * ```
     */
    addAnimation(effect: number, trigger?: number): any;

    /**
     * Removes all animations from the shape
     */
    removeAnimations(): void;

    /**
     * Deletes the shape
     */
    delete(): void;

    /**
     * Releases COM resources
     */
    release(): void;
}

// ============================================================================
// Additional connectors (Access, OneNote, Visio, Acrobat, SAP) would follow
// the same pattern with comprehensive JSDoc comments
// ============================================================================

/**
 * Access Connector - High-level API for Microsoft Access automation
 */
export class AccessConnector {
    constructor();
    openDatabase(path: string): void;
    executeQuery(sql: string): AccessRecordset;
    getRecords(tableName: string): AccessRecordset;
    findRecords(tableName: string, criteria: string): AccessRecordset;
    quit(): void;
    release(): void;
}

/**
 * Access Recordset wrapper for database operations
 */
export class AccessRecordset {
    moveFirst(): void;
    moveLast(): void;
    moveNext(): void;
    movePrevious(): void;
    getFieldValue(fieldName: string): any;
    setFieldValue(fieldName: string, value: any): void;
    addNew(): void;
    update(): void;
    delete(): void;
    getRecordCount(): number;
    isEOF(): boolean;
    isBOF(): boolean;
    findFirst(criteria: string): void;
    findNext(criteria: string): void;
    close(): void;
    release(): void;
}

/**
 * OneNote Connector - High-level API for Microsoft OneNote automation
 */
export class OneNoteConnector {
    constructor();
    openNotebook(path: string): string;
    createNotebook(path: string, name: string): string;
    getNotebooks(): any;
    insertImage(pageId: string, imagePath: string, position?: {x: number, y: number}): void;
    static createFormattedText(text: string, options?: {bold?: boolean, italic?: boolean, fontSize?: number}): string;
    static createTable(rows: number, cols: number): string;
    static createList(items: string[], ordered?: boolean): string;
    quit(): void;
    release(): void;
}

/**
 * Visio Connector - High-level API for Microsoft Visio automation
 */
export class VisioConnector {
    constructor();
    addDocument(): VisioDocument;
    openDocument(path: string): VisioDocument;
    openStencil(path: string): VisioStencil;
    quit(): void;
    release(): void;
}

/**
 * Visio Document wrapper
 */
export class VisioDocument {
    getName(): string;
    getPageCount(): number;
    getPage(index: number): VisioPage;
    addPage(): VisioPage;
    save(): void;
    saveAs(path: string): void;
    close(): void;
    release(): void;
}

/**
 * Visio Page wrapper
 */
export class VisioPage {
    getName(): string;
    setName(name: string): void;
    drawRectangle(x1: number, y1: number, x2: number, y2: number): VisioShape;
    drawLine(x1: number, y1: number, x2: number, y2: number): VisioShape;
    dropMaster(master: any, x: number, y: number): VisioShape;
    connectShapes(fromShape: VisioShape, toShape: VisioShape): VisioShape;
    release(): void;
}

/**
 * Visio Shape wrapper
 */
export class VisioShape {
    getText(): string;
    setText(text: string): void;
    release(): void;
}

/**
 * Visio Stencil wrapper
 */
export class VisioStencil {
    getMasters(): any;
    getMaster(name: string): any;
    getMasterByIndex(index: number): any;
    getMasterCount(): number;
    close(): void;
    release(): void;
}

/**
 * Acrobat PDF-lib Integration for advanced PDF operations
 */
export class AcrobatPDFLib {
    static extractText(pdfPath: string): Promise<string>;
    static mergePDFs(pdfPaths: string[], outputPath: string): Promise<void>;
    static splitPDF(pdfPath: string, outputDir: string, pageRanges: number[][]): Promise<void>;
    static protectPDF(pdfPath: string, outputPath: string, userPassword: string, ownerPassword?: string): Promise<void>;
    static getPageCount(pdfPath: string): Promise<number>;
    static extractPages(pdfPath: string, outputPath: string, pageNumbers: number[]): Promise<void>;
}

/**
 * SAP GUI Connector - High-level API for SAP GUI automation
 */
export class SAPConnector {
    constructor();
    connect(connectionString: string): SAPConnection;
    getScriptingEngine(): any;
    quit(): void;
    release(): void;
}

/**
 * SAP Connection wrapper
 */
export class SAPConnection {
    openConnection(): SAPSession;
    closeConnection(): void;
    release(): void;
}

/**
 * SAP Session wrapper
 */
export class SAPSession {
    findById(id: string): any;
    setFieldValue(id: string, value: string): void;
    getFieldValue(id: string): string;
    pressButton(id: string): void;
    sendVKey(key: number): void;
    getStatusBarText(): string;
    getStatusBarType(): string;
    takeScreenshot(filename: string): void;
    getGridView(id: string): SAPGridView;
    release(): void;
}

/**
 * SAP Grid View wrapper
 */
export class SAPGridView {
    getRowCount(): number;
    getColumnCount(): number;
    getCellValue(row: number, column: string): string;
    setCellValue(row: number, column: string, value: string): void;
    selectRow(row: number): void;
    doubleClickCell(row: number, column: string): void;
    getSelectedRows(): number[];
    pressButton(buttonId: string): void;
    release(): void;
}
