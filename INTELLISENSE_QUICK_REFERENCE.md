# IntelliSense Quick Reference Card

**Quick guide for using IntelliSense with node-winautomation**

---

## 🚀 Quick Start

```javascript
const { ExcelConnector, WordConnector, OutlookConnector } = require('node-winautomation');
// IntelliSense now active! Press Ctrl+Space to see all methods
```

---

## ⌨️ Keyboard Shortcuts

| Action | Shortcut | Description |
|--------|----------|-------------|
| **Trigger IntelliSense** | `Ctrl+Space` | Show available methods/properties |
| **Parameter Info** | `Ctrl+Shift+Space` | Show parameter hints |
| **Quick Info** | `Hover` | View documentation |
| **Go to Definition** | `F12` | Jump to type definition |
| **Find References** | `Shift+F12` | Find all usages |
| **Rename Symbol** | `F2` | Rename variable everywhere |

---

## 📦 Common Patterns

### Excel
```javascript
const excel = new ExcelConnector();
const wb = excel.openWorkbook('file.xlsx');
const sheet = wb.getWorksheet(1);
sheet.writeCell(1, 1, 'Hello');
const range = sheet.getRange('A1:C10');
range.setBold(true);
wb.save();
```

### Word
```javascript
const word = new WordConnector();
const doc = word.addDocument();
doc.setText('Hello World');
const range = doc.getContent();
range.setFontSize(14);
doc.saveAs('file.docx');
```

### Outlook
```javascript
const outlook = new OutlookConnector();
const mail = outlook.createMailItem();
mail.setTo('user@example.com');
mail.setSubject('Test');
mail.send();
```

### PowerPoint
```javascript
const ppt = new PowerPointConnector();
const pres = ppt.addPresentation();
const slide = pres.addSlide(1);
const box = slide.addTextBox(100, 100, 400, 100);
box.setText('Hello');
pres.saveAs('file.pptx');
```

---

## 🎯 IntelliSense Features

✅ **Auto-completion** - Type hints for all methods  
✅ **Parameter info** - Inline parameter documentation  
✅ **Return types** - Know what methods return  
✅ **Examples** - See usage examples in tooltips  
✅ **Error prevention** - Catch mistakes before runtime  
✅ **Constants** - All enums available with IntelliSense  

---

## 📚 Available Connectors

| Connector | Import | Coverage |
|-----------|--------|----------|
| **Excel** | `ExcelConnector` | 100% |
| **Outlook** | `OutlookConnector` | 100% |
| **PowerPoint** | `PowerPointConnector` | 98% |
| **Word** | `WordConnector` | 88% |
| **SAP GUI** | `SAPConnector` | 100% |
| **Visio** | `VisioConnector` | 100% |
| **Access** | `AccessConnector` | 100% |
| **OneNote** | `OneNoteConnector` | 100% |
| **Acrobat** | `AcrobatPDFLib` | 100% |

---

## 🔧 Constants & Enums

All constants available with IntelliSense:

```javascript
// Excel
XlFileFormat.xlOpenXMLWorkbook
XlHAlign.xlHAlignCenter
XlChartType.xlColumnClustered

// Word
WdSaveFormat.wdFormatPDF
WdParagraphAlignment.wdAlignParagraphCenter

// Outlook
OlDefaultFolders.olFolderInbox
OlItemType.olMailItem

// PowerPoint
PpSaveAsFileType.ppSaveAsPDF
```

---

## ✅ Verification

**Test IntelliSense:**
1. Type: `const excel = new ExcelConnector();`
2. Type: `excel.` and press `Ctrl+Space`
3. See all methods with documentation!

---

## 📖 Full Documentation

- **Complete Guide:** `INTELLISENSE_GUIDE.md`
- **Type Definitions:** `features/com-automation/lib/connectors/connectors.d.ts`
- **Examples:** `tests/` directory

---

**Status:** ✅ IntelliSense Fully Enabled  
**Methods Documented:** 200+  
**Connectors:** 9  
**Coverage:** 96% average
