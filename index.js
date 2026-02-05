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
    // Constants
    OlItemType: COMConnectors.OlItemType,
    OlDefaultFolders: COMConnectors.OlDefaultFolders,
    OlImportance: COMConnectors.OlImportance,
    XlFileFormat: COMConnectors.XlFileFormat,
    XlChartType: COMConnectors.XlChartType,
    XlCalculation: COMConnectors.XlCalculation,
};

// Export everything with multiple import styles
module.exports = {
    // Namespaces (for organized imports)
    UIAutomation,
    COMAutomation,
    
    // Direct class exports (standard/professional style)
    ExcelConnector: COMConnectors.ExcelConnector,
    OutlookConnector: COMConnectors.OutlookConnector,
    
    // All other exports for backward compatibility
    ...NativeAddon,
    ...COMConnectors,
};