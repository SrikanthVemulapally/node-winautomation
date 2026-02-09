/**
 * @fileoverview Microsoft Excel Connector
 * Provides a high-level, type-safe API for Excel automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * Excel file formats
 */
const XlFileFormat = {
    xlWorkbookDefault: 51,
    xlOpenXMLWorkbook: 51,
    xlOpenXMLWorkbookMacroEnabled: 52,
    xlExcel8: 56,
    xlCSV: 6,
    xlTextWindows: 20,
    xlHtml: 44
};

/**
 * Chart types
 */
const XlChartType = {
    xlColumnClustered: 51,
    xlColumnStacked: 52,
    xlLine: 4,
    xlLineMarkers: 65,
    xlPie: 5,
    xlBarClustered: 57,
    xlArea: 1,
    xlXYScatter: -4169
};

/**
 * Calculation modes
 */
const XlCalculation = {
    xlCalculationAutomatic: -4105,
    xlCalculationManual: -4135,
    xlCalculationSemiautomatic: 2
};

/**
 * Horizontal alignment constants
 */
const XlHAlign = {
    xlHAlignCenter: -4108,
    xlHAlignLeft: -4131,
    xlHAlignRight: -4152,
    xlHAlignJustify: -4130,
    xlHAlignDistributed: -4117
};

/**
 * Vertical alignment constants
 */
const XlVAlign = {
    xlVAlignCenter: -4108,
    xlVAlignTop: -4160,
    xlVAlignBottom: -4107,
    xlVAlignJustify: -4130,
    xlVAlignDistributed: -4117
};

/**
 * Border position constants
 */
const XlBordersIndex = {
    xlEdgeLeft: 7,
    xlEdgeTop: 8,
    xlEdgeBottom: 9,
    xlEdgeRight: 10,
    xlInsideVertical: 11,
    xlInsideHorizontal: 12,
    xlDiagonalDown: 5,
    xlDiagonalUp: 6
};

/**
 * Line style constants
 */
const XlLineStyle = {
    xlLineStyleNone: -4142,
    xlContinuous: 1,
    xlDash: -4115,
    xlDot: -4118,
    xlDashDot: 4,
    xlDashDotDot: 5,
    xlDouble: -4119
};

/**
 * Border weight constants
 */
const XlBorderWeight = {
    xlHairline: 1,
    xlThin: 2,
    xlMedium: -4138,
    xlThick: 4
};

/**
 * Sort order constants
 */
const XlSortOrder = {
    xlAscending: 1,
    xlDescending: 2
};

/**
 * @class ExcelConnector
 * @description High-level API for Microsoft Excel automation
 */
class ExcelConnector {
    constructor() {
        this.app = new COMObject('Excel.Application');
    }

    /**
     * Get Excel version
     * @returns {string} Version string
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Set application visibility
     * @param {boolean} visible - Visibility flag
     */
    setVisible(visible) {
        this.app.setProperty('Visible', visible);
    }

    /**
     * Get application visibility
     * @returns {boolean} Visibility flag
     */
    getVisible() {
        return this.app.getProperty('Visible');
    }

    /**
     * Set display alerts
     * @param {boolean} display - Display alerts flag
     */
    setDisplayAlerts(display) {
        this.app.setProperty('DisplayAlerts', display);
    }

    /**
     * Set screen updating
     * @param {boolean} updating - Screen updating flag
     */
    setScreenUpdating(updating) {
        this.app.setProperty('ScreenUpdating', updating);
    }

    /**
     * Set calculation mode
     * @param {number} mode - Calculation mode constant
     */
    setCalculation(mode) {
        this.app.setProperty('Calculation', mode);
    }

    /**
     * Get workbooks collection
     * @returns {COMObject} Workbooks collection
     */
    getWorkbooks() {
        return this.app.getProperty('Workbooks');
    }

    /**
     * Add a new workbook
     * @returns {Workbook} Workbook wrapper
     */
    addWorkbook() {
        const workbooks = this.getWorkbooks();
        const workbook = workbooks.invoke('Add');
        return new Workbook(workbook);
    }

    /**
     * Open an existing workbook
     * @param {string} filename - Full path to workbook
     * @param {boolean} readOnly - Open as read-only
     * @returns {Workbook} Workbook wrapper
     */
    openWorkbook(filename, readOnly = false) {
        const workbooks = this.getWorkbooks();
        const workbook = workbooks.invoke('Open', filename, undefined, readOnly);
        return new Workbook(workbook);
    }

    /**
     * Get active workbook
     * @returns {Workbook|null} Active workbook or null
     */
    getActiveWorkbook() {
        try {
            const workbook = this.app.getProperty('ActiveWorkbook');
            return workbook ? new Workbook(workbook) : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * Calculate all workbooks
     */
    calculate() {
        this.app.invoke('Calculate');
    }

    /**
     * Evaluate a formula
     * @param {string} formula - Formula to evaluate
     * @returns {*} Result value
     */
    evaluate(formula) {
        return this.app.invoke('Evaluate', formula);
    }

    /**
     * Quit Excel
     */
    quit() {
        this.app.invoke('Quit');
    }

    /**
     * Release resources
     */
    release() {
        this.app.release();
    }
}

/**
 * @class Workbook
 * @description Wrapper for Excel Workbook
 */
class Workbook {
    constructor(comObject) {
        this.workbook = comObject;
    }

    getName() { return this.workbook.getProperty('Name'); }
    getFullName() { return this.workbook.getProperty('FullName'); }
    getPath() { return this.workbook.getProperty('Path'); }

    getSaved() { return this.workbook.getProperty('Saved'); }
    setSaved(value) { this.workbook.setProperty('Saved', value); }

    getReadOnly() { return this.workbook.getProperty('ReadOnly'); }

    /**
     * Get worksheets collection
     * @returns {COMObject} Worksheets collection
     */
    getWorksheets() {
        return this.workbook.getProperty('Worksheets');
    }

    /**
     * Get worksheet by index (1-based)
     * @param {number} index - Worksheet index
     * @returns {Worksheet} Worksheet wrapper
     */
    getWorksheet(index) {
        const worksheets = this.getWorksheets();
        if (!worksheets) {
            throw new Error('Failed to get worksheets collection');
        }
        const sheet = worksheets.getProperty('Item', index);
        return new Worksheet(sheet);
    }

    /**
     * Get worksheet by name
     * @param {string} name - Worksheet name
     * @returns {Worksheet} Worksheet wrapper
     */
    getWorksheetByName(name) {
        const worksheets = this.getWorksheets();
        if (!worksheets) {
            throw new Error('Failed to get worksheets collection');
        }
        const sheet = worksheets.getProperty('Item', name);
        return new Worksheet(sheet);
    }

    /**
     * Add a new worksheet
     * @param {string} name - Worksheet name (optional)
     * @returns {Worksheet} Worksheet wrapper
     */
    addWorksheet(name = null) {
        const worksheets = this.getWorksheets();
        if (!worksheets) {
            throw new Error('Failed to get worksheets collection');
        }
        const sheet = worksheets.invoke('Add');
        if (name) {
            sheet.setProperty('Name', name);
        }
        return new Worksheet(sheet);
    }

    /**
     * Get active worksheet
     * @returns {Worksheet} Active worksheet
     */
    getActiveSheet() {
        const sheet = this.workbook.getProperty('ActiveSheet');
        return new Worksheet(sheet);
    }

    /**
     * Get worksheet count
     * @returns {number} Number of worksheets
     */
    getWorksheetCount() {
        const worksheets = this.getWorksheets();
        if (!worksheets) {
            throw new Error('Failed to get worksheets collection');
        }
        return worksheets.getProperty('Count');
    }

    /**
     * Save workbook
     */
    save() {
        this.workbook.invoke('Save');
    }

    /**
     * Save workbook as
     * @param {string} filename - Full path
     * @param {number} fileFormat - File format constant
     */
    saveAs(filename, fileFormat = XlFileFormat.xlOpenXMLWorkbook) {
        // Excel SaveAs has many optional parameters, we need to pass them correctly
        // SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, ...)
        this.workbook.invoke('SaveAs', filename, fileFormat, '', '', false, false);
    }

    /**
     * Close workbook
     * @param {boolean} saveChanges - Save changes flag
     */
    close(saveChanges = false) {
        this.workbook.invoke('Close', saveChanges);
    }

    /**
     * Activate workbook
     */
    activate() {
        this.workbook.invoke('Activate');
    }

    /**
     * Export as PDF
     * @param {string} filename - Output filename
     */
    exportAsPDF(filename) {
        this.workbook.invoke('ExportAsFixedFormat', 0, filename);
    }

    release() { this.workbook.release(); }
}

/**
 * @class Worksheet
 * @description Wrapper for Excel Worksheet
 */
class Worksheet {
    constructor(comObject) {
        this.worksheet = comObject;
    }

    getName() { return this.worksheet.getProperty('Name'); }
    setName(value) { this.worksheet.setProperty('Name', value); }

    getIndex() { return this.worksheet.getProperty('Index'); }

    getVisible() { return this.worksheet.getProperty('Visible'); }
    setVisible(value) { this.worksheet.setProperty('Visible', value); }

    /**
     * Get a range by address
     * @param {string} address - Range address (e.g., "A1", "A1:C10")
     * @returns {Range} Range wrapper
     */
    getRange(address) {
        // Excel's Range property is accessed as a parameterized property
        const range = this.worksheet.getProperty('Range', address);
        return new Range(range);
    }

    /**
     * Get a cell by row and column (1-based)
     * @param {number} row - Row number
     * @param {number} column - Column number
     * @returns {Range} Range wrapper for the cell
     */
    getCells(row, column) {
        const cells = this.worksheet.getProperty('Cells');
        const cell = cells.getProperty('Item', row, column);
        return new Range(cell);
    }

    /**
     * Get all cells
     * @returns {Range} Range wrapper for all cells
     */
    getAllCells() {
        const cells = this.worksheet.getProperty('Cells');
        return new Range(cells);
    }

    /**
     * Get used range
     * @returns {Range} Range wrapper for used range
     */
    getUsedRange() {
        const range = this.worksheet.getProperty('UsedRange');
        return new Range(range);
    }

    /**
     * Get last used row
     * @returns {number} Last row number
     */
    getLastRow() {
        const usedRange = this.getUsedRange();
        const rows = usedRange.getRows();
        const rowCount = rows.getProperty('Count');
        const lastRow = usedRange.getRow() + rowCount - 1;
        return lastRow;
    }

    /**
     * Get last used column
     * @returns {number} Last column number
     */
    getLastColumn() {
        const usedRange = this.getUsedRange();
        const columns = usedRange.getColumns();
        const colCount = columns.getProperty('Count');
        const lastCol = usedRange.getColumn() + colCount - 1;
        return lastCol;
    }

    /**
     * Read cell value
     * @param {number} row - Row number (1-based)
     * @param {number} column - Column number (1-based)
     * @returns {*} Cell value
     */
    readCell(row, column) {
        const cells = this.worksheet.getProperty('Cells');
        const cell = cells.getProperty('Item', row, column);
        return cell.getProperty('Value');
    }

    /**
     * Write cell value
     * @param {number} row - Row number (1-based)
     * @param {number} column - Column number (1-based)
     * @param {*} value - Value to write
     */
    writeCell(row, column, value) {
        const cells = this.worksheet.getProperty('Cells');
        const cell = cells.getProperty('Item', row, column);
        cell.setProperty('Value', value);
    }

    /**
     * Read range as 2D array
     * @param {string} address - Range address (e.g., "A1:C10")
     * @returns {Array<Array>} 2D array of values
     */
    readRange(address) {
        const range = this.worksheet.getProperty('Range', address);
        try {
            // Try Value2 first (works better with COM)
            const values = range.getProperty('Value2');
            return values;
        } catch (e) {
            // Fallback to Value
            try {
                const values = range.getProperty('Value');
                return values;
            } catch (e2) {
                // If both fail, return null
                return null;
            }
        }
    }

    /**
     * Write range from 2D array
     * @param {string} address - Range address (e.g., "A1:C10")
     * @param {Array<Array>} data - 2D array of values
     */
    writeRange(address, data) {
        try {
            const range = this.worksheet.getProperty('Range', address);
            range.setProperty('Value', data);
        } catch (e) {
            // Fallback: write cell by cell
            const match = address.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
            if (match) {
                const startRow = parseInt(match[2]);
                const startCol = this._columnLetterToNumber(match[1]);
                for (let r = 0; r < data.length; r++) {
                    for (let c = 0; c < data[r].length; c++) {
                        this.writeCell(startRow + r, startCol + c, data[r][c]);
                    }
                }
            }
        }
    }

    /**
     * Convert column letter to number (A=1, B=2, etc.)
     * @private
     */
    _columnLetterToNumber(letter) {
        let column = 0;
        for (let i = 0; i < letter.length; i++) {
            column = column * 26 + (letter.charCodeAt(i) - 64);
        }
        return column;
    }

    /**
     * Append row to end of used range
     * @param {Array} rowData - Array of values for the row
     */
    appendRow(rowData) {
        try {
            const lastRow = this.getLastRow();
            const nextRow = lastRow + 1;
            for (let i = 0; i < rowData.length; i++) {
                this.writeCell(nextRow, i + 1, rowData[i]);
            }
        } catch (e) {
            // If getLastRow fails (empty sheet), start at row 1
            for (let i = 0; i < rowData.length; i++) {
                this.writeCell(1, i + 1, rowData[i]);
            }
        }
    }

    /**
     * Protect worksheet
     * @param {string} password - Password (optional)
     */
    protect(password = '') {
        if (password) {
            this.worksheet.invoke('Protect', password);
        } else {
            this.worksheet.invoke('Protect');
        }
    }

    /**
     * Unprotect worksheet
     * @param {string} password - Password (optional)
     */
    unprotect(password = '') {
        if (password) {
            this.worksheet.invoke('Unprotect', password);
        } else {
            this.worksheet.invoke('Unprotect');
        }
    }

    /**
     * Set column width
     * @param {number} column - Column number (1-based)
     * @param {number} width - Width value
     */
    setColumnWidth(column, width) {
        const columns = this.worksheet.getProperty('Columns');
        const col = columns.getProperty('Item', column);
        col.setProperty('ColumnWidth', width);
    }

    /**
     * Set row height
     * @param {number} row - Row number (1-based)
     * @param {number} height - Height value
     */
    setRowHeight(row, height) {
        const rows = this.worksheet.getProperty('Rows');
        const rowObj = rows.getProperty('Item', row);
        rowObj.setProperty('RowHeight', height);
    }

    /**
     * Auto fit columns in range
     * @param {string} address - Range address (e.g., "A:C")
     */
    autoFitColumns(address) {
        const range = this.worksheet.getProperty('Range', address);
        const columns = range.getProperty('Columns');
        columns.invoke('AutoFit');
    }

    /**
     * Auto fit rows in range
     * @param {string} address - Range address (e.g., "1:10")
     */
    autoFitRows(address) {
        const range = this.worksheet.getProperty('Range', address);
        const rows = range.getProperty('Rows');
        rows.invoke('AutoFit');
    }

    /**
     * Move worksheet
     * @param {Worksheet} before - Move before this worksheet (optional)
     * @param {Worksheet} after - Move after this worksheet (optional)
     */
    move(before = null, after = null) {
        if (before) {
            this.worksheet.invoke('Move', before.worksheet);
        } else if (after) {
            this.worksheet.invoke('Move', undefined, after.worksheet);
        }
    }

    /**
     * Sort range
     * @param {string} rangeAddress - Range to sort
     * @param {string} keyAddress - Key column/row address
     * @param {number} order - Sort order (1=ascending, 2=descending)
     */
    sortRange(rangeAddress, keyAddress, order = 1) {
        try {
            const range = this.worksheet.getProperty('Range', rangeAddress);
            const key = this.worksheet.getProperty('Range', keyAddress);
            // Use Excel's Sort method with minimal parameters
            range.invoke('Sort', key, order);
        } catch (e) {
            // If Sort fails, try using the worksheet's Sort object
            const range = this.worksheet.getProperty('Range', rangeAddress);
            const sort = this.worksheet.getProperty('Sort');
            const key = this.worksheet.getProperty('Range', keyAddress);
            sort.setProperty('SortFields').invoke('Clear');
            sort.setProperty('SortFields').invoke('Add', key, 0, order);
            sort.setProperty('SetRange', range);
            sort.invoke('Apply');
        }
    }

    /**
     * Activate worksheet
     */
    activate() {
        this.worksheet.invoke('Activate');
    }

    /**
     * Delete worksheet
     */
    delete() {
        this.worksheet.invoke('Delete');
    }

    /**
     * Copy worksheet
     * @param {Worksheet} before - Insert before this worksheet (optional)
     * @param {Worksheet} after - Insert after this worksheet (optional)
     */
    copy(before = null, after = null) {
        if (before) {
            this.worksheet.invoke('Copy', before.worksheet);
        } else if (after) {
            this.worksheet.invoke('Copy', undefined, after.worksheet);
        } else {
            this.worksheet.invoke('Copy');
        }
    }

    /**
     * Calculate worksheet
     */
    calculate() {
        this.worksheet.invoke('Calculate');
    }

    /**
     * Get charts collection
     * @returns {COMObject} Charts collection
     */
    getCharts() {
        return this.worksheet.getProperty('ChartObjects');
    }

    /**
     * Add a chart
     * @returns {Chart} Chart wrapper
     */
    addChart() {
        const charts = this.getCharts();
        if (!charts) {
            throw new Error('Failed to get charts collection');
        }
        const chart = charts.invoke('Add');
        return new Chart(chart);
    }

    /**
     * Add pivot table
     * @param {string} sourceRange - Source data range address
     * @param {string} destinationCell - Destination cell for pivot table
     * @param {string} tableName - Pivot table name
     * @returns {PivotTable} Pivot table wrapper
     */
    addPivotTable(sourceRange, destinationCell, tableName) {
        try {
            const workbook = this.worksheet.getProperty('Parent');
            const pivotCaches = workbook.getProperty('PivotCaches');
            const sourceData = this.worksheet.getProperty('Range', sourceRange);
            const pivotCache = pivotCaches.invoke('Create', 1, sourceData); // xlDatabase
            
            const destRange = this.worksheet.getProperty('Range', destinationCell);
            const pivotTables = this.worksheet.getProperty('PivotTables');
            const pivotTable = pivotTables.invoke('Add', pivotCache, destRange, tableName);
            
            return new PivotTable(pivotTable);
        } catch (e) {
            throw new Error('Failed to create pivot table: ' + e.message);
        }
    }
    
    /**
     * Get pivot table by name
     * @param {string} name - Pivot table name
     * @returns {PivotTable} Pivot table wrapper
     */
    getPivotTable(name) {
        const pivotTables = this.worksheet.getProperty('PivotTables');
        const pivotTable = pivotTables.invoke('Item', name);
        return new PivotTable(pivotTable);
    }
    
    /**
     * Get pivot table count
     * @returns {number} Number of pivot tables
     */
    getPivotTableCount() {
        const pivotTables = this.worksheet.getProperty('PivotTables');
        return pivotTables.getProperty('Count');
    }

    release() { this.worksheet.release(); }
}

/**
 * @class Range
 * @description Wrapper for Excel Range
 */
class Range {
    constructor(comObject) {
        this.range = comObject;
    }

    getValue() { return this.range.getProperty('Value'); }
    setValue(value) { this.range.setProperty('Value', value); }

    getValue2() { return this.range.getProperty('Value2'); }

    getText() { return this.range.getProperty('Text'); }

    getFormula() { return this.range.getProperty('Formula'); }
    setFormula(value) { this.range.setProperty('Formula', value); }

    getFormulaR1C1() { return this.range.getProperty('FormulaR1C1'); }
    setFormulaR1C1(value) { this.range.setProperty('FormulaR1C1', value); }

    getNumberFormat() { return this.range.getProperty('NumberFormat'); }
    setNumberFormat(value) { this.range.setProperty('NumberFormat', value); }

    getAddress() { return this.range.getProperty('Address'); }
    getRow() { return this.range.getProperty('Row'); }
    getColumn() { return this.range.getProperty('Column'); }
    getCount() { return this.range.getProperty('Count'); }

    /**
     * Get rows collection
     * @returns {COMObject} Rows collection
     */
    getRows() {
        return this.range.getProperty('Rows');
    }

    /**
     * Get columns collection
     * @returns {COMObject} Columns collection
     */
    getColumns() {
        return this.range.getProperty('Columns');
    }

    /**
     * Merge cells in range
     */
    merge() {
        this.range.invoke('Merge');
    }

    /**
     * Unmerge cells in range
     */
    unmerge() {
        this.range.invoke('UnMerge');
    }

    /**
     * Get borders
     * @returns {COMObject} Borders collection
     */
    getBorders() {
        return this.range.getProperty('Borders');
    }

    /**
     * Set border style
     * @param {number} position - Border position (7=left, 8=top, 9=bottom, 10=right, 12=all)
     * @param {number} lineStyle - Line style (1=continuous, 0=none)
     * @param {number} weight - Line weight (1=hairline, 2=thin, 3=medium, 4=thick)
     */
    setBorder(position, lineStyle = 1, weight = 2) {
        const borders = this.getBorders();
        const border = borders.getProperty('Item', position);
        border.setProperty('LineStyle', lineStyle);
        border.setProperty('Weight', weight);
    }

    /**
     * Set horizontal alignment
     * @param {number} alignment - Alignment constant (-4108=center, -4131=left, -4152=right)
     */
    setHorizontalAlignment(alignment) {
        this.range.setProperty('HorizontalAlignment', alignment);
    }

    /**
     * Set vertical alignment
     * @param {number} alignment - Alignment constant (-4108=center, -4160=top, -4107=bottom)
     */
    setVerticalAlignment(alignment) {
        this.range.setProperty('VerticalAlignment', alignment);
    }

    /**
     * Set wrap text
     * @param {boolean} wrap - Wrap text flag
     */
    setWrapText(wrap) {
        this.range.setProperty('WrapText', wrap);
    }

    /**
     * Auto fill range
     * @param {string} destinationAddress - Destination range address
     * @param {number} fillType - Fill type (0=default, 1=copy, 2=series, etc.)
     */
    autoFill(destinationAddress, fillType = 0) {
        const worksheet = this.range.getProperty('Worksheet');
        const destination = worksheet.getProperty('Range', destinationAddress);
        this.range.invoke('AutoFill', destination, fillType);
    }

    /**
     * Get row height
     * @returns {number} Row height
     */
    getRowHeight() {
        return this.range.getProperty('RowHeight');
    }

    /**
     * Set row height
     * @param {number} height - Height value
     */
    setRowHeight(height) {
        this.range.setProperty('RowHeight', height);
    }

    /**
     * Get column width
     * @returns {number} Column width
     */
    getColumnWidth() {
        return this.range.getProperty('ColumnWidth');
    }

    /**
     * Set column width
     * @param {number} width - Width value
     */
    setColumnWidth(width) {
        this.range.setProperty('ColumnWidth', width);
    }

    /**
     * Get font object
     * @returns {Font} Font wrapper
     */
    getFont() {
        const font = this.range.getProperty('Font');
        return new Font(font);
    }

    /**
     * Get interior object
     * @returns {Interior} Interior wrapper
     */
    getInterior() {
        const interior = this.range.getProperty('Interior');
        return new Interior(interior);
    }

    /**
     * Select range
     */
    select() {
        this.range.invoke('Select');
    }
    
    /**
     * Add data validation
     * @param {number} type - Validation type (1=WholeNumber, 2=Decimal, 3=List, 4=Date, 5=Time, 6=TextLength, 7=Custom)
     * @param {number} operator - Operator (1=Between, 2=NotBetween, 3=Equal, 4=NotEqual, 5=Greater, 6=Less, 7=GreaterEqual, 8=LessEqual)
     * @param {string} formula1 - First formula/value
     * @param {string} formula2 - Second formula/value (optional)
     */
    addValidation(type, operator, formula1, formula2 = null) {
        const validation = this.range.getProperty('Validation');
        validation.invoke('Delete');
        if (formula2) {
            validation.invoke('Add', type, operator, 1, formula1, formula2);
        } else {
            validation.invoke('Add', type, operator, 1, formula1);
        }
    }
    
    /**
     * Clear data validation
     */
    clearValidation() {
        const validation = this.range.getProperty('Validation');
        validation.invoke('Delete');
    }
    
    /**
     * Add conditional formatting
     * @param {number} type - Format condition type (1=CellValue, 2=Expression)
     * @param {number} operator - Operator (1=Between, 2=NotBetween, 3=Equal, etc.)
     * @param {string} formula1 - First formula/value
     * @param {string} formula2 - Second formula/value (optional)
     * @returns {Object} Format condition object
     */
    addConditionalFormatting(type, operator, formula1, formula2 = null) {
        const formatConditions = this.range.getProperty('FormatConditions');
        let condition;
        if (formula2) {
            condition = formatConditions.invoke('Add', type, operator, formula1, formula2);
        } else {
            condition = formatConditions.invoke('Add', type, operator, formula1);
        }
        return condition;
    }
    
    /**
     * Clear conditional formatting
     */
    clearConditionalFormatting() {
        const formatConditions = this.range.getProperty('FormatConditions');
        formatConditions.invoke('Delete');
    }

    /**
     * Copy range
     * @param {Range} destination - Destination range (optional)
     */
    copy(destination = null) {
        if (destination) {
            this.range.invoke('Copy');
            destination.range.invoke('Select');
            destination.range.getProperty('Worksheet').invoke('Paste');
        } else {
            this.range.invoke('Copy');
        }
    }

    /**
     * Cut range
     * @param {Range} destination - Destination range (optional)
     */
    cut(destination = null) {
        if (destination) {
            this.range.invoke('Cut', destination.range);
        } else {
            this.range.invoke('Cut');
        }
    }

    /**
     * Clear contents and formatting
     */
    clear() {
        this.range.invoke('Clear');
    }

    /**
     * Clear contents only
     */
    clearContents() {
        this.range.invoke('ClearContents');
    }

    /**
     * Clear formatting only
     */
    clearFormats() {
        this.range.invoke('ClearFormats');
    }

    /**
     * Delete cells
     * @param {number} shift - Shift direction
     */
    delete(shift = -4162) { // xlUp
        this.range.invoke('Delete', shift);
    }

    /**
     * Insert cells
     * @param {number} shift - Shift direction
     */
    insert(shift = -4121) { // xlDown
        this.range.invoke('Insert', shift);
    }

    /**
     * Auto fit columns
     */
    autoFit() {
        const columns = this.range.getProperty('Columns');
        if (!columns) {
            throw new Error('Failed to get columns collection');
        }
        columns.invoke('AutoFit');
    }

    /**
     * Find value
     * @param {*} what - Value to find
     * @returns {Range|null} Found range or null
     */
    find(what) {
        try {
            const found = this.range.invoke('Find', what);
            return found ? new Range(found) : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * Replace value
     * @param {*} what - Value to find
     * @param {*} replacement - Replacement value
     * @returns {boolean} Success flag
     */
    replace(what, replacement) {
        try {
            this.range.invoke('Replace', what, replacement);
            return true;
        } catch (e) {
            return false;
        }
    }

    release() { this.range.release(); }
}

/**
 * @class Font
 * @description Wrapper for Excel Font
 */
class Font {
    constructor(comObject) {
        this.font = comObject;
    }

    getName() { return this.font.getProperty('Name'); }
    setName(value) { this.font.setProperty('Name', value); }

    getSize() { return this.font.getProperty('Size'); }
    setSize(value) { this.font.setProperty('Size', value); }

    getBold() { return this.font.getProperty('Bold'); }
    setBold(value) { this.font.setProperty('Bold', value); }

    getItalic() { return this.font.getProperty('Italic'); }
    setItalic(value) { this.font.setProperty('Italic', value); }

    getUnderline() { return this.font.getProperty('Underline'); }
    setUnderline(value) { this.font.setProperty('Underline', value); }

    getColor() { return this.font.getProperty('Color'); }
    setColor(value) { this.font.setProperty('Color', value); }

    release() { this.font.release(); }
}

/**
 * @class Interior
 * @description Wrapper for Excel Interior (cell background)
 */
class Interior {
    constructor(comObject) {
        this.interior = comObject;
    }

    getColor() { return this.interior.getProperty('Color'); }
    setColor(value) { this.interior.setProperty('Color', value); }

    getColorIndex() { return this.interior.getProperty('ColorIndex'); }
    setColorIndex(value) { this.interior.setProperty('ColorIndex', value); }

    getPattern() { return this.interior.getProperty('Pattern'); }
    setPattern(value) { this.interior.setProperty('Pattern', value); }

    release() { this.interior.release(); }
}

/**
 * @class Chart
 * @description Wrapper for Excel Chart
 */
class Chart {
    constructor(comObject) {
        this.chart = comObject;
    }

    getChartType() { return this.chart.getProperty('ChartType'); }
    setChartType(value) { this.chart.setProperty('ChartType', value); }

    getHasTitle() { return this.chart.getProperty('HasTitle'); }
    setHasTitle(value) { this.chart.setProperty('HasTitle', value); }

    getHasLegend() { return this.chart.getProperty('HasLegend'); }
    setHasLegend(value) { this.chart.setProperty('HasLegend', value); }

    /**
     * Get chart title
     * @returns {COMObject} ChartTitle object
     */
    getChartTitle() {
        return this.chart.getProperty('ChartTitle');
    }

    /**
     * Set chart title text
     * @param {string} text - Title text
     */
    setTitleText(text) {
        this.setHasTitle(true);
        const title = this.getChartTitle();
        title.setProperty('Text', text);
    }

    /**
     * Set data source
     * @param {Range} source - Source range
     */
    setSourceData(source) {
        this.chart.invoke('SetSourceData', source.range);
    }

    /**
     * Export chart
     * @param {string} filename - Output filename
     */
    export(filename) {
        this.chart.invoke('Export', filename);
    }

    /**
     * Delete chart
     */
    delete() {
        this.chart.invoke('Delete');
    }

    release() { this.chart.release(); }
}

/**
 * @class PivotTable
 * @description Wrapper for Excel Pivot Table
 */
class PivotTable {
    constructor(comObject) {
        this.pivotTable = comObject;
    }
    
    /**
     * Get pivot table name
     * @returns {string} Pivot table name
     */
    getName() {
        return this.pivotTable.getProperty('Name');
    }
    
    /**
     * Set pivot table name
     * @param {string} name - New name
     */
    setName(name) {
        this.pivotTable.setProperty('Name', name);
    }
    
    /**
     * Add row field
     * @param {string} fieldName - Field name from source data
     * @param {number} position - Position in row area (optional)
     */
    addRowField(fieldName, position = null) {
        const pivotFields = this.pivotTable.getProperty('PivotFields');
        const field = pivotFields.invoke('Item', fieldName);
        field.setProperty('Orientation', 1); // xlRowField
        if (position !== null) {
            field.setProperty('Position', position);
        }
    }
    
    /**
     * Add column field
     * @param {string} fieldName - Field name from source data
     * @param {number} position - Position in column area (optional)
     */
    addColumnField(fieldName, position = null) {
        const pivotFields = this.pivotTable.getProperty('PivotFields');
        const field = pivotFields.invoke('Item', fieldName);
        field.setProperty('Orientation', 2); // xlColumnField
        if (position !== null) {
            field.setProperty('Position', position);
        }
    }
    
    /**
     * Add data field
     * @param {string} fieldName - Field name from source data
     * @param {number} function - Aggregation function (default: -4157 = xlSum)
     * @param {string} caption - Custom caption (optional)
     */
    addDataField(fieldName, functionType = -4157, caption = null) {
        const pivotFields = this.pivotTable.getProperty('PivotFields');
        const field = pivotFields.invoke('Item', fieldName);
        field.setProperty('Orientation', 4); // xlDataField
        field.setProperty('Function', functionType);
        if (caption) {
            field.setProperty('Caption', caption);
        }
    }
    
    /**
     * Add page field (filter)
     * @param {string} fieldName - Field name from source data
     * @param {number} position - Position in page area (optional)
     */
    addPageField(fieldName, position = null) {
        const pivotFields = this.pivotTable.getProperty('PivotFields');
        const field = pivotFields.invoke('Item', fieldName);
        field.setProperty('Orientation', 3); // xlPageField
        if (position !== null) {
            field.setProperty('Position', position);
        }
    }
    
    /**
     * Refresh pivot table data
     */
    refresh() {
        this.pivotTable.invoke('RefreshTable');
    }
    
    /**
     * Clear pivot table
     */
    clear() {
        this.pivotTable.invoke('ClearTable');
    }
    
    /**
     * Release COM object
     */
    release() {
        this.pivotTable.release();
    }
}

module.exports = {
    ExcelConnector,
    Workbook,
    Worksheet,
    Range,
    Font,
    Interior,
    Chart,
    PivotTable,
    XlFileFormat,
    XlChartType,
    XlCalculation,
    XlHAlign,
    XlVAlign,
    XlBordersIndex,
    XlLineStyle,
    XlBorderWeight,
    XlSortOrder,
};
