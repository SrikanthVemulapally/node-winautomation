/**
 * COM Automation Connectors
 * High-level APIs for Microsoft Office automation
 */

const OutlookConnector = require('./OutlookConnector');
const ExcelConnector = require('./ExcelConnector');
const WordConnector = require('./WordConnector');
const PowerPointConnector = require('./PowerPointConnector');
const AcrobatConnector = require('./AcrobatConnector');
const AccessConnector = require('./AccessConnector');
const OneNoteConnector = require('./OneNoteConnector');
const VisioConnector = require('./VisioConnector');
const SAPConnector = require('./SAPConnector');

module.exports = {
    // Outlook
    OutlookConnector: OutlookConnector.OutlookConnector,
    MailItem: OutlookConnector.MailItem,
    AppointmentItem: OutlookConnector.AppointmentItem,
    ContactItem: OutlookConnector.ContactItem,
    TaskItem: OutlookConnector.TaskItem,
    Folder: OutlookConnector.Folder,
    OlItemType: OutlookConnector.OlItemType,
    OlDefaultFolders: OutlookConnector.OlDefaultFolders,
    OlImportance: OutlookConnector.OlImportance,
    
    // Excel
    ExcelConnector: ExcelConnector.ExcelConnector,
    Workbook: ExcelConnector.Workbook,
    Worksheet: ExcelConnector.Worksheet,
    Range: ExcelConnector.Range,
    Font: ExcelConnector.Font,
    Interior: ExcelConnector.Interior,
    Chart: ExcelConnector.Chart,
    XlFileFormat: ExcelConnector.XlFileFormat,
    XlChartType: ExcelConnector.XlChartType,
    XlCalculation: ExcelConnector.XlCalculation,
    XlHAlign: ExcelConnector.XlHAlign,
    XlVAlign: ExcelConnector.XlVAlign,
    XlBordersIndex: ExcelConnector.XlBordersIndex,
    XlLineStyle: ExcelConnector.XlLineStyle,
    XlBorderWeight: ExcelConnector.XlBorderWeight,
    XlSortOrder: ExcelConnector.XlSortOrder,
    
    // Word
    WordConnector: WordConnector.WordConnector,
    WordDocument: WordConnector.WordDocument,
    WordRange: WordConnector.WordRange,
    WordParagraph: WordConnector.WordParagraph,
    WordParagraphFormat: WordConnector.WordParagraphFormat,
    WordTable: WordConnector.WordTable,
    WordFont: WordConnector.WordFont,
    WordPageSetup: WordConnector.WordPageSetup,
    WdSaveFormat: WordConnector.WdSaveFormat,
    WdParagraphAlignment: WordConnector.WdParagraphAlignment,
    WdOrientation: WordConnector.WdOrientation,
    WdUnderline: WordConnector.WdUnderline,
    WdColor: WordConnector.WdColor,
    WdBreakType: WordConnector.WdBreakType,
    
    // PowerPoint
    PowerPointConnector: PowerPointConnector.PowerPointConnector,
    Presentation: PowerPointConnector.Presentation,
    Slide: PowerPointConnector.Slide,
    Shape: PowerPointConnector.Shape,
    TextFrame: PowerPointConnector.TextFrame,
    TextRange: PowerPointConnector.TextRange,
    PPFont: PowerPointConnector.PPFont,
    ParagraphFormat: PowerPointConnector.ParagraphFormat,
    PPTable: PowerPointConnector.PPTable,
    PpSaveAsFileType: PowerPointConnector.PpSaveAsFileType,
    PpSlideLayout: PowerPointConnector.PpSlideLayout,
    MsoShapeType: PowerPointConnector.MsoShapeType,
    MsoAutoShapeType: PowerPointConnector.MsoAutoShapeType,
    PpParagraphAlignment: PowerPointConnector.PpParagraphAlignment,
    
    // Acrobat
    AcrobatConnector: AcrobatConnector.AcrobatConnector,
    PDFDocument: AcrobatConnector.PDFDocument,
    
    // Access
    AccessConnector: AccessConnector.AccessConnector,
    
    // OneNote
    OneNoteConnector: OneNoteConnector.OneNoteConnector,
    
    // Visio
    VisioConnector: VisioConnector.VisioConnector,
    VisioDocument: VisioConnector.VisioDocument,
    VisioPage: VisioConnector.VisioPage,
    VisioShape: VisioConnector.VisioShape,
    VisSaveAsFileType: VisioConnector.VisSaveAsFileType,
    
    // SAP
    SAPConnector: SAPConnector.SAPConnector,
    SAPApplication: SAPConnector.SAPApplication,
    SAPConnection: SAPConnector.SAPConnection,
    SAPSession: SAPConnector.SAPSession,
    SAPGridView: SAPConnector.SAPGridView,
    SAPVKey: SAPConnector.SAPVKey,
    
    // Additional classes
    PivotTable: ExcelConnector.PivotTable,
    AccessRecordset: AccessConnector.AccessRecordset,
    AcrobatPDFLib: require('./AcrobatPDFLib').AcrobatPDFLib,
};
