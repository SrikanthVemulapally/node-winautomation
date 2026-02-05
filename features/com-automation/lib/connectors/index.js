/**
 * COM Automation Connectors
 * High-level APIs for Microsoft Office automation
 */

const OutlookConnector = require('./OutlookConnector');
const ExcelConnector = require('./ExcelConnector');

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
    XlCalculation: ExcelConnector.XlCalculation
};
