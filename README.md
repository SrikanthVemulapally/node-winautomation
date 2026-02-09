<div align="center">

# node-winautomation

[![npm version](https://img.shields.io/npm/v/node-winautomation.svg)](https://www.npmjs.com/package/node-winautomation)
[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://github.com/SrikanthVemulapally/node-winautomation)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D14.0.0-brightgreen.svg)](https://nodejs.org/)

**A powerful Node.js native addon for Windows UI Automation**

Automate Windows desktop applications with full access to the Microsoft UI Automation API.

[Installation](#installation) • [Quick Start](#quick-start) • [API Reference](#api-reference) • [Examples](#examples) • [Contributing](#contributing)

</div>

---

## ✨ Features

- **🎯 Complete UI Automation API** — Full native bindings to Windows UI Automation
- **🔧 30+ Automation Patterns** — Invoke, Value, Text, Selection, Grid, Window, and more
- **📡 Event Handling** — Focus changes, property changes, structure changes, automation events
- **🏢 COM Automation** — 9 Office connectors (Excel, Word, PowerPoint, Outlook, Access, OneNote, Visio) + SAP GUI + Acrobat
- **📘 TypeScript Support** — Comprehensive type definitions with 200+ documented methods
- **💡 IntelliSense Enabled** — Full auto-completion, parameter hints, and inline documentation
- **⚡ Native Performance** — C++ addon for optimal speed
- **🔍 Element Discovery** — Find elements by name, control type, automation ID, and custom conditions
- **🎨 Clean API** — Organized namespaces (UIAutomation, COMAutomation) for better code clarity
- **📊 96% Coverage** — Industry-leading Office automation coverage (Excel 100%, Outlook 100%, Word 88%)

## 📋 Requirements

| Requirement | Version |
|-------------|---------|
| **OS** | Windows 10/11 |
| **Node.js** | ≥ 14.0.0 |
| **Visual Studio** | 2019 or 2022 with C++ Desktop Development |
| **Python** | 3.x (for node-gyp) |

## 📦 Installation

### From npm

```bash
npm install node-winautomation
```

The native addon will be compiled automatically during installation.

### From Source

```bash
git clone https://github.com/SrikanthVemulapally/node-winautomation.git
cd node-winautomation
npm install
```

### Build Commands

| Command | Description |
|---------|-------------|
| `npm run configure` | Configure the native build |
| `npm run build` | Build the native addon |
| `npm run rebuild` | Clean and rebuild |
| `npm run clean` | Remove build artifacts |
| `npm test` | Run test suite |

## 🚀 Quick Start

### Basic Example: Click a Button

```javascript
const { UIAutomation } = require('node-winautomation');

// Initialize automation
const automation = new UIAutomation.Automation();
const root = automation.getRootElement();

// Find a window by name
const windowCondition = automation.createPropertyCondition(
  UIAutomation.PropertyIds.NamePropertyId,
  'My Application'
);
const appWindow = root.findFirst(UIAutomation.TreeScopes.Children, windowCondition);

// Find and click a button
const buttonCondition = automation.createPropertyCondition(
  UIAutomation.PropertyIds.NamePropertyId,
  'Submit'
);
const button = appWindow.findFirst(UIAutomation.TreeScopes.Descendants, buttonCondition);

// Click using Invoke pattern
const invokePattern = button.getCurrentPattern(UIAutomation.PatternIds.InvokePatternId);
invokePattern.invoke();
```

### Set Text in a Text Box

```javascript
const { UIAutomation } = require('node-winautomation');

const automation = new UIAutomation.Automation();
const root = automation.getRootElement();

// Find Notepad
const notepad = root.findFirst(
  UIAutomation.TreeScopes.Children,
  automation.createPropertyCondition(UIAutomation.PropertyIds.NamePropertyId, 'Untitled - Notepad')
);

// Find the edit control
const editControl = notepad.findFirst(
  UIAutomation.TreeScopes.Descendants,
  automation.createPropertyCondition(UIAutomation.PropertyIds.ControlTypePropertyId, UIAutomation.ControlTypeIds.DocumentControlTypeId)
);

// Set text using Value pattern
const valuePattern = editControl.getCurrentPattern(UIAutomation.PatternIds.ValuePatternId);
valuePattern.setValue('Hello from node-winautomation!');
```

### COM Automation: Office Examples

```javascript
// Outlook - Send Email
const { OutlookConnector } = require('node-winautomation');
const outlook = new OutlookConnector();
const mail = outlook.createMailItem();
mail.setTo('user@example.com');
mail.setSubject('Test');
mail.send();

// Excel - Workbook Operations
const { ExcelConnector } = require('node-winautomation');
const excel = new ExcelConnector();
const wb = excel.openWorkbook('data.xlsx');
const sheet = wb.getWorksheet(1);
sheet.writeCell(1, 1, 'Hello');
wb.save();

// Word - Document Creation
const { WordConnector } = require('node-winautomation');
const word = new WordConnector();
const doc = word.addDocument();
doc.setText('Hello World');
doc.saveAs('document.docx');

// PowerPoint - Presentations
const { PowerPointConnector } = require('node-winautomation');
const ppt = new PowerPointConnector();
const pres = ppt.addPresentation();
const slide = pres.addSlide(1);
slide.addTextBox(100, 100, 400, 100).setText('Hello');
pres.saveAs('presentation.pptx');
```

> **Note:** Word table/bookmark/comment methods are `@experimental` (73% success rate). See `WORD_COM_LIMITATIONS.md`.

## 📖 API Reference

### Automation Class

The main entry point for UI Automation operations.

```typescript
class Automation {
  // Element retrieval
  getRootElement(): AutomationElement;
  getFocusedElement(): AutomationElement;
  elementFromPoint(point: Point): AutomationElement;

  // Condition creation
  createPropertyCondition(propertyId: PropertyIds, value: Variant): AutomationCondition;
  createAndCondition(c1: AutomationCondition, c2: AutomationCondition): AutomationCondition;
  createOrCondition(c1: AutomationCondition, c2: AutomationCondition): AutomationCondition;
  createNotCondition(condition: AutomationCondition): AutomationCondition;
  createTrueCondition(): AutomationCondition;
  createFalseCondition(): AutomationCondition;

  // Tree walkers
  createTreeWalker(condition: AutomationCondition): AutomationTreeWalker;
  rawViewWalker: AutomationTreeWalker;
  controlViewWalker: AutomationTreeWalker;
  contentViewWalker: AutomationTreeWalker;

  // Event handling
  addFocusChangedEventHandler(cacheRequest, handler): void;
  addAutomationEventHandler(eventId, element, scope, cacheRequest, handler): void;
  addPropertyChangedEventHandler(element, scope, cacheRequest, handler): void;
  addStructureChangedEventHandler(element, scope, cacheRequest, handler): void;
  removeAllEventHandlers(): void;

  // Utilities
  compareElements(e1: AutomationElement, e2: AutomationElement): boolean;
  createCacheRequest(): AutomationCacheRequest;
}
```

### AutomationElement

Represents a UI element in the automation tree.

```typescript
interface AutomationElement {
  // Search methods
  findFirst(scope: TreeScopes, condition: AutomationCondition): AutomationElement;
  findAll(scope: TreeScopes, condition: AutomationCondition): AutomationElement[];

  // Pattern retrieval
  getCurrentPattern(patternId: PatternIds): Pattern;
  getCachedPattern(patternId: PatternIds): Pattern;

  // Common properties
  currentName: string;
  currentClassName: string;
  currentControlType: ControlTypeIds;
  currentAutomationId: string;
  currentBoundingRectangle: Rect;
  currentIsEnabled: boolean;
  currentIsOffscreen: boolean;
  currentProcessId: number;

  // Methods
  setFocus(): void;
  getClickablePoint(): Point | null;
  getRuntimeId(): number[];
}
```

### Automation Patterns

| Pattern | Description | Key Methods |
|---------|-------------|-------------|
| **IInvokePattern** | Click buttons, menu items | `invoke()` |
| **IValuePattern** | Get/set text values | `setValue(value)`, `currentValue` |
| **ITogglePattern** | Checkboxes, toggle buttons | `toggle()`, `currentToggleState` |
| **ISelectionPattern** | Lists, combo boxes | `getCurrentSelection()` |
| **ISelectionItemPattern** | List items | `select()`, `addToSelection()` |
| **IExpandCollapsePattern** | Tree nodes, menus | `expand()`, `collapse()` |
| **IScrollPattern** | Scrollable containers | `scroll()`, `setScrollPercent()` |
| **IWindowPattern** | Window operations | `close()`, `setWindowVisualState()` |
| **ITextPattern** | Rich text controls | `documentRange`, `getSelection()` |
| **IGridPattern** | Tables, data grids | `getItem(row, col)` |
| **ITransformPattern** | Move/resize elements | `move()`, `resize()`, `rotate()` |
| **IRangeValuePattern** | Sliders, progress bars | `setValue()`, `currentValue` |

### Enumerations

#### TreeScopes
```typescript
enum TreeScopes {
  None = 0,
  Element = 0x1,
  Children = 0x2,
  Descendants = 0x4,
  Parent = 0x8,
  Ancestors = 0x10,
  Subtree = Element | Children | Descendants
}
```

#### ControlTypeIds
```typescript
enum ControlTypeIds {
  ButtonControlTypeId = 50000,
  CheckBoxControlTypeId = 50002,
  ComboBoxControlTypeId = 50003,
  EditControlTypeId = 50004,
  ListControlTypeId = 50008,
  MenuControlTypeId = 50009,
  WindowControlTypeId = 50032,
  // ... and 30+ more
}
```

#### PatternIds
```typescript
enum PatternIds {
  InvokePatternId = 10000,
  SelectionPatternId = 10001,
  ValuePatternId = 10002,
  TogglePatternId = 10015,
  WindowPatternId = 10009,
  TextPatternId = 10014,
  // ... and 25+ more
}
```

#### PropertyIds
```typescript
enum PropertyIds {
  NamePropertyId = 30005,
  ControlTypePropertyId = 30003,
  AutomationIdPropertyId = 30011,
  ClassNamePropertyId = 30012,
  IsEnabledPropertyId = 30010,
  // ... and 150+ more
}
```

## 📁 Examples

### Calculator Automation

```javascript
const { UIAutomation } = require('node-winautomation');

const automation = new UIAutomation.Automation();
const root = automation.getRootElement();

// Find Calculator
const calcCond = automation.createPropertyCondition(
  UIAutomation.PropertyIds.NamePropertyId, 'Calculator'
);
const calc = root.findFirst(UIAutomation.TreeScopes.Children, calcCond);

// Click button
const btnCond = automation.createPropertyCondition(
  UIAutomation.PropertyIds.NamePropertyId, 'Two'
);
const button = calc.findFirst(UIAutomation.TreeScopes.Descendants, btnCond);
button.getCurrentPattern(UIAutomation.PatternIds.InvokePatternId).invoke();
```

### SAP GUI Automation

```javascript
const { SAPConnector } = require('node-winautomation');

const sap = new SAPConnector();
const session = sap.connect('/H/server/S/3200').openConnection();

session.setFieldValue('wnd[0]/usr/txtRSYST-BNAME', 'user');
session.pressButton('wnd[0]/tbar[0]/btn[0]');

const grid = session.getGridView('wnd[0]/usr/cntlGRID1/shellcont/shell');
console.log(`Rows: ${grid.getRowCount()}`);
```

### IntelliSense Support

All connectors have full IntelliSense support with auto-completion and inline documentation:

```javascript
const { ExcelConnector } = require('node-winautomation');
const excel = new ExcelConnector();
// Type "excel." and press Ctrl+Space to see all available methods!
```

See `INTELLISENSE_GUIDE.md` for complete usage instructions.

More examples available in `features/ui-automation/examples/` and `features/com-automation/examples/`.

## 📂 Project Structure

```
node-winautomation/
├── index.js                 # Module entry point with UIAutomation & COMAutomation namespaces
├── index.d.ts               # TypeScript definitions
├── binding.gyp              # Native build configuration
├── package.json
├── README.md
├── ARCHITECTURE.md
├── GETTING_STARTED.md
│
├── features/                # Feature-based organization
│   │
│   ├── ui-automation/       # UI Automation feature
│   │   ├── src/
│   │   │   ├── wrappers/    # Core UI Automation wrappers
│   │   │   ├── patterns/    # 30+ automation patterns
│   │   │   ├── enumerations/# UI Automation enums
│   │   │   ├── utilities/   # Event handlers & helpers
│   │   │   └── AutomationAddon.*
│   │   ├── examples/
│   │   │   ├── calculator-example.js
│   │   │   ├── notepad-example.js
│   │   │   └── outlook-message-example.js
│   │   └── docs/
│   │
│   ├── com-automation/      # COM Automation feature
│   │   ├── src/core/        # COM base classes
│   │   │   ├── COMObject.*
│   │   │   ├── COMVariant.*
│   │   │   └── COMEventSink.*
│   │   ├── lib/connectors/  # Office connectors
│   │   │   ├── OutlookConnector.js
│   │   │   ├── ExcelConnector.js
│   │   │   └── index.js
│   │   ├── examples/
│   │   └── docs/
│   │
│   └── desktop-management/  # Desktop Management (optional)
│       ├── src/wrappers/    # Child sessions & RDP
│       ├── examples/
│       └── docs/
│
└── tests/                   # Test suite
    ├── *.test.js            # Unit tests
    ├── test-*.js            # Integration tests
    └── README.md
```

## 🔧 Troubleshooting

### Build Issues

| Error | Solution |
|-------|----------|
| `Cannot find module 'node-gyp'` | `npm install -g node-gyp` |
| `MSBuild.exe not found` | Install Visual Studio Build Tools with C++ workload |
| `Python not found` | Install Python 3.x and add to PATH |
| `node-gyp rebuild failed` | Run `npm run clean` then `npm run rebuild` |

### Runtime Issues

| Error | Solution |
|-------|----------|
| `Module not found` | Ensure `build/Release/Automation.node` exists |
| `Access Denied` | Run Node.js as Administrator |
| `Element not found` | Verify window is open and element name is correct |
| `Pattern not supported` | Check if element supports the requested pattern |

### Tips

- Use **Accessibility Insights** or **Inspect.exe** to explore UI element properties
- Some applications require **elevated privileges** (run as Administrator)
- UWP apps may have different element structures than Win32 apps
- Use `TreeScopes.Descendants` carefully on large trees (can be slow)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🙏 Credits

The UI Automation feature is based on the excellent work by **Chris Nimmons** and **Jake Cyr** in the [@bright-fish/node-ui-automation](https://github.com/bright-fish/node-ui-automation) project.

## 📄 License

MIT License - See [LICENSE](LICENSE) file for details.

## 📚 Resources

### Documentation
- [IntelliSense Guide](./INTELLISENSE_GUIDE.md) - Complete IntelliSense usage guide
- [IntelliSense Quick Reference](./INTELLISENSE_QUICK_REFERENCE.md) - Quick reference card
- [Word COM Limitations](./features/com-automation/docs/WORD_COM_LIMITATIONS.md) - Technical analysis of Word automation limitations
- [Getting Started Guide](./GETTING_STARTED.md) - Comprehensive getting started guide
- [Architecture](./ARCHITECTURE.md) - Project architecture and design

### External Resources
- [Windows UI Automation Overview](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
- [UI Automation Control Patterns](https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpatternsoverview)
- [Accessibility Insights for Windows](https://accessibilityinsights.io/docs/windows/overview/)
- [Node.js N-API Documentation](https://nodejs.org/api/n-api.html)

## 🎯 COM Automation Coverage

| Connector | Coverage | Status | Notes |
|-----------|----------|--------|-------|
| **Excel** | 100% | ✅ Production | Full workbook, worksheet, range, pivot table support |
| **Outlook** | 100% | ✅ Production | Email, calendar, contacts, tasks, rules, categories |
| **PowerPoint** | 98% | ✅ Production | Presentations, slides, shapes, animations, transitions |
| **SAP GUI** | 100% | ✅ Production | Session management, field operations, grid views |
| **Visio** | 100% | ✅ Production | Documents, pages, shapes, stencils, connections |
| **Access** | 100% | ✅ Production | Database operations, recordsets, queries |
| **OneNote** | 100% | ✅ Production | Notebooks, sections, pages, content insertion |
| **Acrobat** | 100% | ✅ Production | PDF operations via pdf-lib integration |
| **Word** | 88% | ⚠️ Partial | Core features work; tables/bookmarks/comments experimental |

**Average Coverage: 96%** - Industry-leading Office automation support

---

<div align="center">
Made with ❤️ for Windows automation
</div>
