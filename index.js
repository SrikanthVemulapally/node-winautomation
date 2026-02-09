/**
 * @fileoverview node-winautomation - Windows Automation Library
 * @description Complete Windows automation solution with UI Automation and COM support
 * 
 * Provides comprehensive automation for:
 * - Microsoft Office (Excel, Word, PowerPoint, Outlook, Access, OneNote, Visio)
 * - SAP GUI
 * - Adobe Acrobat
 * - Windows UI Automation
 * 
 * @example
 * // Excel automation
 * const { ExcelConnector } = require('node-winautomation');
 * const excel = new ExcelConnector();
 * const workbook = excel.openWorkbook('C:\\data.xlsx');
 * 
 * @example
 * // UI Automation
 * const { UIAutomation } = require('node-winautomation');
 * const automation = new UIAutomation.Automation();
 * const element = automation.getRootElement();
 * 
 * @version 1.0.0
 * @license MIT
 */

/**
 * node-winautomation
 * Windows automation library with UI Automation and COM support
 */

// Load native addon
const NativeAddon = require('bindings')('Automation');

// Load COM automation connectors
const COMConnectors = require('./features/com-automation/lib');

// Create organized namespaces
const UIAutomation = {
    Automation: NativeAddon.Automation,
    PropertyIds: NativeAddon.PropertyIds,
    ControlTypeIds: NativeAddon.ControlTypeIds,
    PatternIds: NativeAddon.PatternIds,
    TreeScopes: NativeAddon.TreeScopes,
    EventIds: NativeAddon.EventIds,
    ToggleStates: NativeAddon.ToggleStates,
    WindowVisualStates: NativeAddon.WindowVisualStates,
    WindowInteractionStates: NativeAddon.WindowInteractionStates,
    DockPositions: NativeAddon.DockPositions,
    ExpandCollapseStates: NativeAddon.ExpandCollapseStates,
    ScrollAmounts: NativeAddon.ScrollAmounts,
    OrientationTypes: NativeAddon.OrientationTypes,
    // Add other enumerations as needed
};

const COMAutomation = {
    COMObject: NativeAddon.COMObject,
    OutlookConnector: COMConnectors.OutlookConnector,
    ExcelConnector: COMConnectors.ExcelConnector,
    WordConnector: COMConnectors.WordConnector,
    PowerPointConnector: COMConnectors.PowerPointConnector,
    AcrobatConnector: COMConnectors.AcrobatConnector,
    AccessConnector: COMConnectors.AccessConnector,
    OneNoteConnector: COMConnectors.OneNoteConnector,
    VisioConnector: COMConnectors.VisioConnector,
    SAPConnector: COMConnectors.SAPConnector,
    // Outlook classes
    MailItem: COMConnectors.MailItem,
    AppointmentItem: COMConnectors.AppointmentItem,
    ContactItem: COMConnectors.ContactItem,
    TaskItem: COMConnectors.TaskItem,
    Folder: COMConnectors.Folder,
    // Excel classes
    Workbook: COMConnectors.Workbook,
    Worksheet: COMConnectors.Worksheet,
    Range: COMConnectors.Range,
    Font: COMConnectors.Font,
    Interior: COMConnectors.Interior,
    Chart: COMConnectors.Chart,
    // Word classes
    WordDocument: COMConnectors.WordDocument,
    WordRange: COMConnectors.WordRange,
    WordParagraph: COMConnectors.WordParagraph,
    WordParagraphFormat: COMConnectors.WordParagraphFormat,
    WordTable: COMConnectors.WordTable,
    WordFont: COMConnectors.WordFont,
    WordPageSetup: COMConnectors.WordPageSetup,
    // PowerPoint classes
    Presentation: COMConnectors.Presentation,
    Slide: COMConnectors.Slide,
    Shape: COMConnectors.Shape,
    TextFrame: COMConnectors.TextFrame,
    TextRange: COMConnectors.TextRange,
    PPFont: COMConnectors.PPFont,
    ParagraphFormat: COMConnectors.ParagraphFormat,
    PPTable: COMConnectors.PPTable,
    // Constants
    OlItemType: COMConnectors.OlItemType,
    OlDefaultFolders: COMConnectors.OlDefaultFolders,
    OlImportance: COMConnectors.OlImportance,
    XlFileFormat: COMConnectors.XlFileFormat,
    XlChartType: COMConnectors.XlChartType,
    XlCalculation: COMConnectors.XlCalculation,
    XlHAlign: COMConnectors.XlHAlign,
    XlVAlign: COMConnectors.XlVAlign,
    XlBordersIndex: COMConnectors.XlBordersIndex,
    XlLineStyle: COMConnectors.XlLineStyle,
    XlBorderWeight: COMConnectors.XlBorderWeight,
    XlSortOrder: COMConnectors.XlSortOrder,
    WdSaveFormat: COMConnectors.WdSaveFormat,
    WdParagraphAlignment: COMConnectors.WdParagraphAlignment,
    WdOrientation: COMConnectors.WdOrientation,
    WdUnderline: COMConnectors.WdUnderline,
    WdColor: COMConnectors.WdColor,
    WdBreakType: COMConnectors.WdBreakType,
    PpSaveAsFileType: COMConnectors.PpSaveAsFileType,
    PpSlideLayout: COMConnectors.PpSlideLayout,
    MsoShapeType: COMConnectors.MsoShapeType,
    MsoAutoShapeType: COMConnectors.MsoAutoShapeType,
    PpParagraphAlignment: COMConnectors.PpParagraphAlignment,
};

// Export everything with multiple import styles
module.exports = {
    // Namespaces (for organized imports)
    UIAutomation,
    COMAutomation,
    
    // Direct class exports (standard/professional style)
    ExcelConnector: COMConnectors.ExcelConnector,
    OutlookConnector: COMConnectors.OutlookConnector,
    WordConnector: COMConnectors.WordConnector,
    PowerPointConnector: COMConnectors.PowerPointConnector,
    AcrobatConnector: COMConnectors.AcrobatConnector,
    AccessConnector: COMConnectors.AccessConnector,
    OneNoteConnector: COMConnectors.OneNoteConnector,
    VisioConnector: COMConnectors.VisioConnector,
    SAPConnector: COMConnectors.SAPConnector,
    
    // All other exports for backward compatibility
    ...NativeAddon,
    ...COMConnectors,
};