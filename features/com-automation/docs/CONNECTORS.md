# Office Connectors Guide

## Overview

The Office Connectors provide high-level, type-safe APIs for automating Microsoft Office applications. Built on top of the robust COM automation system, these connectors offer:

- **Type-safe interfaces** - JavaScript classes with clear method signatures
- **Memory management** - Automatic resource cleanup
- **Error handling** - Comprehensive error messages
- **Event support** - Subscribe to Office application events
- **Full API coverage** - Access to all major Office features

## Available Connectors

### OutlookConnector
Automate Microsoft Outlook for email, calendar, contacts, and tasks.

### ExcelConnector
Automate Microsoft Excel for spreadsheet operations, charts, and data analysis.

## OutlookConnector

### Installation

```javascript
const { OutlookConnector } = require('node-winautomation/lib/connectors/OutlookConnector');
```

### Basic Usage

#### Creating and Sending Email

```javascript
const outlook = new OutlookConnector();

// Create email
const mail = outlook.createMailItem();
mail.setTo('recipient@example.com');
mail.setCC('cc@example.com');
mail.setSubject('Hello from Node.js');
mail.setBody('This email was sent using OutlookConnector!');
mail.setImportance(OlImportance.olImportanceHigh);

// Add attachment
mail.addAttachment('C:\\documents\\report.pdf');

// Send or save to drafts
mail.send();  // Send immediately
// mail.save();  // Save to drafts
```

#### Creating Appointments

```javascript
const appointment = outlook.createAppointmentItem();

const start = new Date('2024-01-15 10:00');
const end = new Date('2024-01-15 11:00');

appointment.setSubject('Team Meeting');
appointment.setLocation('Conference Room A');
appointment.setStart(start);
appointment.setEnd(end);
appointment.setBody('Quarterly review meeting');
appointment.setReminderSet(true);
appointment.setReminderMinutesBeforeStart(15);

appointment.save();
```

#### Managing Contacts

```javascript
const contact = outlook.createContactItem();

contact.setFirstName('John');
contact.setLastName('Doe');
contact.setEmail1Address('john.doe@example.com');
contact.setBusinessTelephoneNumber('+1-555-0123');
contact.setMobileTelephoneNumber('+1-555-0124');
contact.setCompanyName('Acme Corp');
contact.setJobTitle('Software Engineer');

contact.save();
```

#### Creating Tasks

```javascript
const task = outlook.createTaskItem();

task.setSubject('Complete project documentation');
task.setBody('Write comprehensive docs for the new feature');
task.setDueDate(new Date('2024-01-20'));
task.setPriority(OlImportance.olImportanceHigh);
task.setPercentComplete(25);

task.save();
```

#### Accessing Folders

```javascript
// Get default folders
const inbox = outlook.getInbox();
const sentMail = outlook.getSentMail();
const drafts = outlook.getDrafts();
const calendar = outlook.getCalendar();
const contacts = outlook.getContacts();

// Read folder properties
console.log('Inbox name:', inbox.getName());
console.log('Unread count:', inbox.getUnReadItemCount());
console.log('Total items:', inbox.getCount());

// Iterate through items
for (let i = 1; i <= inbox.getCount(); i++) {
    const item = inbox.getItem(i);
    console.log('Subject:', item.getProperty('Subject'));
    item.release();
}
```

#### Event Handling

```javascript
// Subscribe to new mail event
outlook.onNewMail((entryId) => {
    console.log('New mail received:', entryId);
    
    // Get the mail item
    const ns = outlook.getNamespace();
    const mail = ns.invoke('GetItemFromID', entryId);
    console.log('From:', mail.getProperty('SenderName'));
    console.log('Subject:', mail.getProperty('Subject'));
    mail.release();
});

// Subscribe to item send event
outlook.onItemSend((item, cancel) => {
    console.log('Sending email:', item.getProperty('Subject'));
    // Set cancel to true to prevent sending
});

// Remove all event subscriptions
outlook.removeAllEvents();
```

### API Reference

#### OutlookConnector Class

**Methods:**
- `getVersion()` - Get Outlook version
- `createMailItem()` - Create new email
- `createAppointmentItem()` - Create new appointment
- `createContactItem()` - Create new contact
- `createTaskItem()` - Create new task
- `getDefaultFolder(folderType)` - Get folder by type
- `getInbox()` - Get inbox folder
- `getSentMail()` - Get sent mail folder
- `getDrafts()` - Get drafts folder
- `getCalendar()` - Get calendar folder
- `getContacts()` - Get contacts folder
- `getTasks()` - Get tasks folder
- `onNewMail(callback)` - Subscribe to new mail event
- `onItemSend(callback)` - Subscribe to item send event
- `removeAllEvents()` - Unsubscribe from all events
- `quit()` - Close Outlook
- `release()` - Release resources

#### MailItem Class

**Properties (via get/set methods):**
- `Subject`, `Body`, `HTMLBody`
- `To`, `CC`, `BCC`
- `Importance`, `Sensitivity`
- `SentOn`, `ReceivedTime`
- `UnRead`, `Size`

**Methods:**
- `send()` - Send email
- `display(modal)` - Display email window
- `save()` - Save to drafts
- `reply()` - Create reply
- `replyAll()` - Create reply all
- `forward()` - Create forward
- `addAttachment(path, type, position, displayName)` - Add attachment

#### AppointmentItem Class

**Properties:**
- `Subject`, `Location`, `Body`
- `Start`, `End`
- `AllDayEvent`
- `ReminderSet`, `ReminderMinutesBeforeStart`
- `BusyStatus`

**Methods:**
- `save()`, `send()`, `display(modal)`

#### ContactItem Class

**Properties:**
- `FirstName`, `LastName`, `FullName`
- `Email1Address`, `Email2Address`, `Email3Address`
- `BusinessTelephoneNumber`, `HomeTelephoneNumber`, `MobileTelephoneNumber`
- `CompanyName`, `JobTitle`

**Methods:**
- `save()`, `display(modal)`, `delete()`

#### TaskItem Class

**Properties:**
- `Subject`, `Body`
- `StartDate`, `DueDate`
- `Status`, `Priority`
- `PercentComplete`, `Complete`

**Methods:**
- `save()`, `display(modal)`

## ExcelConnector

### Installation

```javascript
const { ExcelConnector, XlFileFormat, XlChartType } = require('node-winautomation/lib/connectors/ExcelConnector');
```

### Basic Usage

#### Creating Workbooks

```javascript
const excel = new ExcelConnector();
excel.setVisible(true);
excel.setDisplayAlerts(false);

// Create new workbook
const workbook = excel.addWorkbook();

// Open existing workbook
const workbook2 = excel.openWorkbook('C:\\data\\report.xlsx');

// Get active workbook
const activeWb = excel.getActiveWorkbook();
```

#### Working with Worksheets

```javascript
// Get worksheet
const sheet = workbook.getActiveSheet();
const sheet2 = workbook.getWorksheet(1);  // By index (1-based)
const sheet3 = workbook.getWorksheetByName('Data');  // By name

// Add new worksheet
const newSheet = workbook.addWorksheet('Summary');

// Rename worksheet
sheet.setName('Sales Data');

// Get worksheet count
const count = workbook.getWorksheetCount();
```

#### Writing Data

```javascript
// Single cell
sheet.getCells(1, 1).setValue('Name');
sheet.getCells(1, 2).setValue('Age');

// Using range
const range = sheet.getRange('A1:C1');
range.setValue(['Name', 'Age', 'City']);

// Write multiple rows
const data = [
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles'],
    ['Charlie', 35, 'Chicago']
];

for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
        sheet.getCells(i + 2, j + 1).setValue(data[i][j]);
    }
}
```

#### Reading Data

```javascript
// Single cell
const value = sheet.getCells(1, 1).getValue();

// Range
const range = sheet.getRange('A1:C10');
const values = range.getValue();  // Returns 2D array

// Used range
const usedRange = sheet.getUsedRange();
const allData = usedRange.getValue();
```

#### Formatting

```javascript
// Font formatting
const range = sheet.getRange('A1:C1');
const font = range.getFont();
font.setBold(true);
font.setSize(14);
font.setColor(0xFF0000);  // Red

// Background color
const interior = range.getInterior();
interior.setColor(0xFFFF00);  // Yellow

// Number formatting
const salaryRange = sheet.getRange('D2:D10');
salaryRange.setNumberFormat('$#,##0.00');

// Auto-fit columns
range.autoFit();
```

#### Formulas

```javascript
// Simple formula
sheet.getCells(1, 5).setFormula('=SUM(A1:D1)');

// R1C1 style
sheet.getCells(2, 5).setFormulaR1C1('=SUM(RC[-4]:RC[-1])');

// Array formula
const range = sheet.getRange('E1:E10');
range.setFormula('=A1:A10 * B1:B10');
```

#### Charts

```javascript
// Create chart
const chart = sheet.addChart();

// Set data source
const dataRange = sheet.getRange('A1:C10');
chart.setSourceData(dataRange);

// Configure chart
chart.setChartType(XlChartType.xlColumnClustered);
chart.setHasTitle(true);
chart.setTitleText('Sales by Region');
chart.setHasLegend(true);

// Export chart
chart.export('C:\\charts\\sales.png');
```

#### Find and Replace

```javascript
// Find
const allCells = sheet.getAllCells();
const found = allCells.find('SearchText');
if (found) {
    console.log('Found at:', found.getAddress());
    found.setValue('NewValue');
}

// Replace
allCells.replace('OldText', 'NewText');
```

#### Saving

```javascript
// Save
workbook.save();

// Save as
workbook.saveAs('C:\\data\\output.xlsx', XlFileFormat.xlOpenXMLWorkbook);

// Save as CSV
workbook.saveAs('C:\\data\\output.csv', XlFileFormat.xlCSV);

// Export as PDF
workbook.exportAsPDF('C:\\data\\output.pdf');
```

### API Reference

#### ExcelConnector Class

**Methods:**
- `getVersion()` - Get Excel version
- `setVisible(visible)` - Set application visibility
- `setDisplayAlerts(display)` - Set display alerts
- `setScreenUpdating(updating)` - Set screen updating
- `setCalculation(mode)` - Set calculation mode
- `addWorkbook()` - Create new workbook
- `openWorkbook(filename, readOnly)` - Open workbook
- `getActiveWorkbook()` - Get active workbook
- `calculate()` - Calculate all workbooks
- `evaluate(formula)` - Evaluate formula
- `quit()` - Close Excel
- `release()` - Release resources

#### Workbook Class

**Methods:**
- `getName()`, `getFullName()`, `getPath()`
- `getSaved()`, `setSaved(value)`
- `getWorksheet(index)` - Get by index
- `getWorksheetByName(name)` - Get by name
- `addWorksheet(name)` - Add new worksheet
- `getActiveSheet()` - Get active worksheet
- `getWorksheetCount()` - Get count
- `save()` - Save workbook
- `saveAs(filename, fileFormat)` - Save as
- `close(saveChanges)` - Close workbook
- `exportAsPDF(filename)` - Export as PDF

#### Worksheet Class

**Methods:**
- `getName()`, `setName(value)`
- `getRange(address)` - Get range by address
- `getCells(row, column)` - Get cell
- `getAllCells()` - Get all cells
- `getUsedRange()` - Get used range
- `activate()` - Activate worksheet
- `delete()` - Delete worksheet
- `copy(before, after)` - Copy worksheet
- `calculate()` - Calculate worksheet
- `addChart()` - Add chart

#### Range Class

**Methods:**
- `getValue()`, `setValue(value)`
- `getFormula()`, `setFormula(value)`
- `getNumberFormat()`, `setNumberFormat(value)`
- `getAddress()`, `getRow()`, `getColumn()`
- `getFont()` - Get font object
- `getInterior()` - Get interior object
- `select()` - Select range
- `copy(destination)` - Copy range
- `clear()`, `clearContents()`, `clearFormats()`
- `delete(shift)`, `insert(shift)`
- `autoFit()` - Auto-fit columns
- `find(what)` - Find value
- `replace(what, replacement)` - Replace value

## Constants

### Outlook Constants

```javascript
const { OlItemType, OlDefaultFolders, OlImportance } = require('.../OutlookConnector');

// Item types
OlItemType.olMailItem = 0
OlItemType.olAppointmentItem = 1
OlItemType.olContactItem = 2
OlItemType.olTaskItem = 3

// Default folders
OlDefaultFolders.olFolderInbox = 6
OlDefaultFolders.olFolderSentMail = 5
OlDefaultFolders.olFolderDrafts = 16
OlDefaultFolders.olFolderCalendar = 9
OlDefaultFolders.olFolderContacts = 10
OlDefaultFolders.olFolderTasks = 13

// Importance
OlImportance.olImportanceLow = 0
OlImportance.olImportanceNormal = 1
OlImportance.olImportanceHigh = 2
```

### Excel Constants

```javascript
const { XlFileFormat, XlChartType, XlCalculation } = require('.../ExcelConnector');

// File formats
XlFileFormat.xlOpenXMLWorkbook = 51  // .xlsx
XlFileFormat.xlExcel8 = 56  // .xls
XlFileFormat.xlCSV = 6  // .csv

// Chart types
XlChartType.xlColumnClustered = 51
XlChartType.xlLine = 4
XlChartType.xlPie = 5
XlChartType.xlXYScatter = -4169

// Calculation modes
XlCalculation.xlCalculationAutomatic = -4105
XlCalculation.xlCalculationManual = -4135
```

## Best Practices

### Memory Management

Always release resources when done:

```javascript
const outlook = new OutlookConnector();
try {
    const mail = outlook.createMailItem();
    // ... use mail ...
    mail.release();
} finally {
    outlook.release();
}
```

### Error Handling

Wrap operations in try-catch:

```javascript
try {
    const workbook = excel.openWorkbook('file.xlsx');
    // ... work with workbook ...
} catch (error) {
    console.error('Failed to open workbook:', error.message);
}
```

### Performance

For bulk operations, disable screen updating:

```javascript
excel.setScreenUpdating(false);
excel.setCalculation(XlCalculation.xlCalculationManual);

// ... bulk operations ...

excel.setCalculation(XlCalculation.xlCalculationAutomatic);
excel.setScreenUpdating(true);
```

### Event Cleanup

Always remove event subscriptions:

```javascript
outlook.onNewMail(callback);

// When done
outlook.removeAllEvents();
outlook.release();
```

## Examples

See the `examples/com/` directory for comprehensive examples:

- `outlook-comprehensive-test.js` - Full Outlook API demonstration
- `excel-comprehensive-test.js` - Full Excel API demonstration

## Troubleshooting

### Office Not Installed

**Error:** "Failed to get CLSID from ProgID"

**Solution:** Ensure Microsoft Office is installed

### Permission Issues

**Error:** "Access is denied"

**Solution:** Run as Administrator or check DCOM permissions

### Events Not Firing

**Problem:** Event callbacks not called

**Solution:** Keep Node.js process alive with `setInterval` or event loop

## See Also

- [COM API Reference](./COM_API.md)
- [Memory Management Guide](./MEMORY_MANAGEMENT.md)
- [Outlook COM API Research](./OUTLOOK_COM_API_RESEARCH.md)
- [Excel COM API Research](./EXCEL_COM_API_RESEARCH.md)
