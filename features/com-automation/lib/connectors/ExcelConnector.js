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
        const sheet = worksheets.invoke('Item', index);
        return new Worksheet(sheet);
    }

    /**
     * Get worksheet by name
     * @param {string} name - Worksheet name
     * @returns {Worksheet} Worksheet wrapper
     */
    getWorksheetByName(name) {
        const worksheets = this.getWorksheets();
        const sheet = worksheets.invoke('Item', name);
        return new Worksheet(sheet);
    }

    /**
     * Add a new worksheet
     * @param {string} name - Worksheet name (optional)
     * @returns {Worksheet} Worksheet wrapper
     */
    addWorksheet(name = null) {
        const worksheets = this.getWorksheets();
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
        const cell = cells.invoke('Item', row, column);
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
        return this.worksheet.getProperty('Charts');
    }

    /**
     * Add a chart
     * @returns {Chart} Chart wrapper
     */
    addChart() {
        const charts = this.getCharts();
        const chart = charts.invoke('Add');
        return new Chart(chart);
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
     * Copy range
     * @param {Range} destination - Destination range (optional)
     */
    copy(destination = null) {
        if (destination) {
            this.range.invoke('Copy', destination.range);
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

module.exports = {
    ExcelConnector,
    Workbook,
    Worksheet,
    Range,
    Font,
    Interior,
    Chart,
    XlFileFormat,
    XlChartType,
    XlCalculation
};
