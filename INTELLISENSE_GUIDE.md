# IntelliSense Usage Guide

This guide shows how to use the comprehensive IntelliSense support for all COM automation connectors.

---

## 🎯 IntelliSense Features

### ✅ What You Get:

1. **Auto-completion** - Type hints for all methods and properties
2. **Parameter information** - Inline documentation for parameters
3. **Return type information** - Know what each method returns
4. **Usage examples** - See examples directly in your IDE
5. **Error prevention** - Catch mistakes before runtime
6. **Quick documentation** - Hover over methods for full docs

---

## 📝 Setup for IntelliSense

### VS Code Setup

1. **Install the package:**
   ```bash
   npm install node-winautomation
   ```

2. **Import with JSDoc type hints:**
   ```javascript
   // Method 1: Destructuring with JSDoc
   const { ExcelConnector, Workbook, Worksheet } = require('node-winautomation');
   
   // Method 2: Full import with type annotation
   /** @type {import('node-winautomation')} */
   const automation = require('node-winautomation');
   ```

3. **Enable IntelliSense in VS Code:**
   - IntelliSense is automatically enabled
   - Press `Ctrl+Space` to trigger suggestions
   - Hover over methods to see documentation

### TypeScript Setup

1. **Create tsconfig.json:**
   ```json
   {
     "compilerOptions": {
       "target": "ES2020",
       "module": "commonjs",
       "allowJs": true,
       "checkJs": true,
       "esModuleInterop": true
     }
   }
   ```

2. **Use with TypeScript:**
   ```typescript
   import { ExcelConnector, XlFileFormat } from 'node-winautomation';
   
   const excel = new ExcelConnector();
   // Full IntelliSense support!
   ```

---

## 💡 Usage Examples with IntelliSense

### Excel Automation

```javascript
const { ExcelConnector, XlFileFormat, XlHAlign } = require('node-winautomation');

// IntelliSense shows: constructor()
const excel = new ExcelConnector();

// IntelliSense shows: setVisible(visible: boolean): void
excel.setVisible(true);

// IntelliSense shows: openWorkbook(path: string, readOnly?: boolean): Workbook
const workbook = excel.openWorkbook('C:\\data.xlsx');

// IntelliSense shows: getWorksheet(index: number): Worksheet
const sheet = workbook.getWorksheet(1);

// IntelliSense shows: writeCell(row: number, col: number, value: string | number): void
sheet.writeCell(1, 1, 'Hello');

// IntelliSense shows: getRange(address: string): Range
const range = sheet.getRange('A1:C10');

// IntelliSense shows: setBold(bold: boolean): void
range.setBold(true);

// IntelliSense shows: setHorizontalAlignment(alignment: XlHAlign): void
range.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

// IntelliSense shows: addPivotTable(sourceRange: string, destinationCell: string, tableName: string): PivotTable
const pivot = sheet.addPivotTable('A1:D100', 'F1', 'SalesPivot');

// IntelliSense shows: addRowField(fieldName: string, position?: number): void
pivot.addRowField('Product');

// IntelliSense shows: addDataField(fieldName: string, functionType?: number, caption?: string): void
pivot.addDataField('Sales', -4157, 'Total Sales'); // -4157 = xlSum

// IntelliSense shows: refresh(): void
pivot.refresh();

// IntelliSense shows: saveAs(path: string, format?: XlFileFormat): void
workbook.saveAs('C:\\output.xlsx', XlFileFormat.xlOpenXMLWorkbook);

// IntelliSense shows: close(saveChanges?: boolean): void
workbook.close();

// IntelliSense shows: quit(saveChanges?: boolean): void
excel.quit();
```

### Word Automation

```javascript
const { WordConnector, WdSaveFormat, WdParagraphAlignment } = require('node-winautomation');

// IntelliSense shows all available methods
const word = new WordConnector();
word.setVisible(false);

// IntelliSense shows: addDocument(): WordDocument
const doc = word.addDocument();

// IntelliSense shows: setText(text: string): void
doc.setText('Hello World');

// IntelliSense shows: addParagraph(text: string): WordParagraph
const para = doc.addParagraph('New paragraph');

// IntelliSense shows: setAlignment(alignment: WdParagraphAlignment): void
para.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);

// IntelliSense shows: getContent(): WordRange
const range = doc.getContent();

// IntelliSense shows: setBold(bold: boolean): void
range.setBold(true);

// IntelliSense shows: setFontSize(size: number): void
range.setFontSize(14);

// IntelliSense shows: find(text: string): boolean
const found = doc.find('Hello');

// IntelliSense shows: replace(findText: string, replaceText: string, replaceAll?: boolean): number
const count = doc.replace('Hello', 'Hi', true);

// IntelliSense shows: exportToPDF(path: string): void
doc.exportToPDF('C:\\document.pdf');

// IntelliSense shows: saveAs(path: string, format?: WdSaveFormat): void
doc.saveAs('C:\\document.docx', WdSaveFormat.wdFormatDocumentDefault);

doc.close();
word.quit();
```

### Outlook Automation

```javascript
const { OutlookConnector, OlDefaultFolders } = require('node-winautomation');

const outlook = new OutlookConnector();

// IntelliSense shows: createMailItem(): MailItem
const mail = outlook.createMailItem();

// IntelliSense shows: setTo(to: string): void
mail.setTo('user@example.com');

// IntelliSense shows: setSubject(subject: string): void
mail.setSubject('Test Email');

// IntelliSense shows: setBody(body: string): void
mail.setBody('Hello from Node.js!');

// IntelliSense shows: addCategory(category: string): void
mail.addCategory('Important');

// IntelliSense shows: setVotingOptions(options: string): void
mail.setVotingOptions('Yes;No;Maybe');

// IntelliSense shows: send(): void
mail.send();

// IntelliSense shows: createAppointmentItem(): AppointmentItem
const appt = outlook.createAppointmentItem();

// IntelliSense shows: setSubject(subject: string): void
appt.setSubject('Meeting');

// IntelliSense shows: setStart(start: Date): void
appt.setStart(new Date());

// IntelliSense shows: setDuration(minutes: number): void
appt.setDuration(60);

// IntelliSense shows: save(): void
appt.save();

// IntelliSense shows: getDefaultFolder(folderType: OlDefaultFolders): Folder
const inbox = outlook.getDefaultFolder(OlDefaultFolders.olFolderInbox);

// IntelliSense shows: getCount(): number
const count = inbox.getCount();

// IntelliSense shows: getUnReadItemCount(): number
const unread = inbox.getUnReadItemCount();
```

### PowerPoint Automation

```javascript
const { PowerPointConnector, PpSaveAsFileType } = require('node-winautomation');

const ppt = new PowerPointConnector();
ppt.setVisible(true);

// IntelliSense shows: addPresentation(): Presentation
const pres = ppt.addPresentation();

// IntelliSense shows: addSlide(layout: number): Slide
const slide = pres.addSlide(1); // 1 = blank layout

// IntelliSense shows: addTextBox(left: number, top: number, width: number, height: number): Shape
const textbox = slide.addTextBox(100, 100, 400, 100);

// IntelliSense shows: setText(text: string): void
textbox.setText('Hello PowerPoint!');

// IntelliSense shows: setTransition(type: number, speed?: number): void
slide.setTransition(257, 2); // 257 = Fade, 2 = Medium speed

// IntelliSense shows: addAnimation(effect: number, trigger?: number): any
textbox.addAnimation(10, 0); // 10 = Fade, 0 = On click

// IntelliSense shows: setNotes(text: string): void
slide.setNotes('Speaker notes here');

// IntelliSense shows: saveAs(path: string, format?: PpSaveAsFileType): void
pres.saveAs('C:\\presentation.pptx', PpSaveAsFileType.ppSaveAsOpenXMLPresentation);

pres.close();
ppt.quit();
```

### Access Database Automation

```javascript
const { AccessConnector } = require('node-winautomation');

const access = new AccessConnector();

// IntelliSense shows: openDatabase(path: string): void
access.openDatabase('C:\\database.accdb');

// IntelliSense shows: executeQuery(sql: string): AccessRecordset
const rs = access.executeQuery('SELECT * FROM Customers WHERE City = "London"');

// IntelliSense shows: moveFirst(): void
rs.moveFirst();

// IntelliSense shows: getFieldValue(fieldName: string): any
while (!rs.isEOF()) {
    const name = rs.getFieldValue('CustomerName');
    const email = rs.getFieldValue('Email');
    console.log(`${name}: ${email}`);
    
    // IntelliSense shows: moveNext(): void
    rs.moveNext();
}

// IntelliSense shows: close(): void
rs.close();

// IntelliSense shows: getRecords(tableName: string): AccessRecordset
const allRecords = access.getRecords('Products');

// IntelliSense shows: addNew(): void
allRecords.addNew();

// IntelliSense shows: setFieldValue(fieldName: string, value: any): void
allRecords.setFieldValue('ProductName', 'New Product');
allRecords.setFieldValue('Price', 29.99);

// IntelliSense shows: update(): void
allRecords.update();

allRecords.close();
access.quit();
```

### SAP GUI Automation

```javascript
const { SAPConnector } = require('node-winautomation');

const sap = new SAPConnector();

// IntelliSense shows: connect(connectionString: string): SAPConnection
const conn = sap.connect('/H/sapserver/S/3200');

// IntelliSense shows: openConnection(): SAPSession
const session = conn.openConnection();

// IntelliSense shows: setFieldValue(id: string, value: string): void
session.setFieldValue('wnd[0]/usr/txtRSYST-BNAME', 'username');
session.setFieldValue('wnd[0]/usr/pwdRSYST-BCODE', 'password');

// IntelliSense shows: pressButton(id: string): void
session.pressButton('wnd[0]/tbar[0]/btn[0]');

// IntelliSense shows: getFieldValue(id: string): string
const value = session.getFieldValue('wnd[0]/usr/txtMaterial');

// IntelliSense shows: getGridView(id: string): SAPGridView
const grid = session.getGridView('wnd[0]/usr/cntlGRID1/shellcont/shell');

// IntelliSense shows: getRowCount(): number
const rowCount = grid.getRowCount();

// IntelliSense shows: getCellValue(row: number, column: string): string
for (let i = 0; i < rowCount; i++) {
    const material = grid.getCellValue(i, 'MATERIAL');
    const quantity = grid.getCellValue(i, 'QUANTITY');
    console.log(`${material}: ${quantity}`);
}

// IntelliSense shows: takeScreenshot(filename: string): void
session.takeScreenshot('C:\\sap_screen.png');

session.release();
conn.closeConnection();
sap.quit();
```

---

## 🔍 IntelliSense Tips

### 1. **Trigger IntelliSense Manually**
Press `Ctrl+Space` at any time to see available methods and properties.

### 2. **View Parameter Info**
Press `Ctrl+Shift+Space` inside function parentheses to see parameter hints.

### 3. **Quick Info**
Hover over any method or property to see full documentation.

### 4. **Go to Definition**
Press `F12` on any method to jump to its type definition.

### 5. **Find All References**
Press `Shift+F12` to find all usages of a method.

### 6. **Rename Symbol**
Press `F2` to rename a variable and update all references.

---

## 📚 Constants and Enums

All constants are available with IntelliSense:

```javascript
// Excel constants
XlFileFormat.xlOpenXMLWorkbook
XlChartType.xlColumnClustered
XlHAlign.xlHAlignCenter
XlVAlign.xlVAlignTop

// Word constants
WdSaveFormat.wdFormatPDF
WdParagraphAlignment.wdAlignParagraphCenter

// Outlook constants
OlItemType.olMailItem
OlDefaultFolders.olFolderInbox

// PowerPoint constants
PpSaveAsFileType.ppSaveAsPDF
```

---

## ⚠️ Error Prevention

IntelliSense helps prevent common errors:

```javascript
// ❌ Wrong - IntelliSense will show error
sheet.writeCell('A1', 'value'); // Wrong parameter types

// ✅ Correct - IntelliSense guides you
sheet.writeCell(1, 1, 'value'); // Correct: row, col, value

// ❌ Wrong - Method doesn't exist
workbook.saveTo('file.xlsx'); // IntelliSense won't suggest this

// ✅ Correct - IntelliSense shows available methods
workbook.saveAs('file.xlsx'); // Correct method name
```

---

## 🎓 Learning with IntelliSense

### Discover Available Methods

```javascript
const excel = new ExcelConnector();
const workbook = excel.addWorkbook();
const sheet = workbook.getWorksheet(1);

// Type "sheet." and press Ctrl+Space to see all available methods:
sheet. // IntelliSense shows:
       // - writeCell()
       // - readCell()
       // - writeRange()
       // - readRange()
       // - getRange()
       // - addPivotTable()
       // - getLastRow()
       // - getLastColumn()
       // - appendRow()
       // - delete()
       // - copy()
       // - release()
       // ... and more!
```

### Explore Method Parameters

```javascript
// Type the method name and opening parenthesis
sheet.writeRange( // IntelliSense shows:
                  // writeRange(range: string, data: any[][]): void
                  // Parameter 1: range - Range address (e.g., "A1:C3")
                  // Parameter 2: data - 2D array of values
```

---

## 📖 Additional Resources

- **Type Definitions:** `features/com-automation/lib/connectors/connectors.d.ts`
- **Examples:** `tests/` directory
- **Documentation:** `features/com-automation/docs/`
- **API Reference:** See individual connector files

---

## ✅ Verification

To verify IntelliSense is working:

1. Create a new `.js` file
2. Type: `const { ExcelConnector } = require('node-winautomation');`
3. Type: `const excel = new ExcelConnector();`
4. Type: `excel.` and press `Ctrl+Space`
5. You should see all available methods with documentation!

---

**IntelliSense Status:** ✅ **Fully Enabled**

All 9 COM connectors have comprehensive IntelliSense support with:
- 200+ documented methods
- Parameter type information
- Return type information
- Usage examples
- Error prevention
- Quick documentation
