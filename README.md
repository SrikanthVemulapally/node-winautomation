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
- **📘 TypeScript Support** — Comprehensive type definitions included
- **⚡ Native Performance** — C++ addon for optimal speed
- **🔍 Element Discovery** — Find elements by name, control type, automation ID, and custom conditions

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
const { Automation, PropertyIds, TreeScopes, PatternIds } = require('node-winautomation');

// Initialize automation
const automation = new Automation();
const root = automation.getRootElement();

// Find a window by name
const windowCondition = automation.createPropertyCondition(
  PropertyIds.NamePropertyId,
  'My Application'
);
const appWindow = root.findFirst(TreeScopes.Children, windowCondition);

// Find and click a button
const buttonCondition = automation.createPropertyCondition(
  PropertyIds.NamePropertyId,
  'Submit'
);
const button = appWindow.findFirst(TreeScopes.Descendants, buttonCondition);

// Click using Invoke pattern
const invokePattern = button.getCurrentPattern(PatternIds.InvokePatternId);
invokePattern.invoke();
```

### Set Text in a Text Box

```javascript
const { Automation, PropertyIds, ControlTypeIds, TreeScopes, PatternIds } = require('node-winautomation');

const automation = new Automation();
const root = automation.getRootElement();

// Find Notepad
const notepad = root.findFirst(
  TreeScopes.Children,
  automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Untitled - Notepad')
);

// Find the edit control
const editControl = notepad.findFirst(
  TreeScopes.Descendants,
  automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.EditControlTypeId)
);

// Set text using Value pattern
const valuePattern = editControl.getCurrentPattern(PatternIds.ValuePatternId);
valuePattern.setValue('Hello from node-winautomation!');
```

### Listen for Focus Changes

```javascript
const { 
  Automation, 
  AutomationFocusChangedEventHandler 
} = require('node-winautomation');

const automation = new Automation();

// Create event handler
const focusHandler = new AutomationFocusChangedEventHandler((element) => {
  console.log('Focus changed to:', element.currentName);
});

// Register handler
automation.addFocusChangedEventHandler(null, focusHandler);

// Keep process running
setInterval(() => {}, 1000);

// Cleanup when done
// automation.removeFocusChangedEventHandler(focusHandler);
```

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
// examples/calculator-example.js
const { Automation, PropertyIds, TreeScopes, PatternIds } = require('node-winautomation');

async function clickButton(window, buttonName, automation) {
  const condition = automation.createPropertyCondition(PropertyIds.NamePropertyId, buttonName);
  const button = window.findFirst(TreeScopes.Descendants, condition);
  
  if (button) {
    const invokePattern = button.getCurrentPattern(PatternIds.InvokePatternId);
    invokePattern.invoke();
    await new Promise(r => setTimeout(r, 200));
  }
}

async function calculate() {
  const automation = new Automation();
  const root = automation.getRootElement();
  
  // Find Calculator
  const calcCondition = automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Calculator');
  const calculator = root.findFirst(TreeScopes.Children, calcCondition);
  
  if (calculator) {
    // Calculate 2 + 3 = 5
    await clickButton(calculator, 'Two', automation);
    await clickButton(calculator, 'Plus', automation);
    await clickButton(calculator, 'Three', automation);
    await clickButton(calculator, 'Equals', automation);
  }
}

calculate();
```

### Element Inspector

```javascript
const { Automation, TreeScopes } = require('node-winautomation');

function inspectElement(element, depth = 0) {
  const indent = '  '.repeat(depth);
  console.log(`${indent}Name: ${element.currentName}`);
  console.log(`${indent}Type: ${element.currentLocalizedControlType}`);
  console.log(`${indent}Class: ${element.currentClassName}`);
  console.log(`${indent}AutomationId: ${element.currentAutomationId}`);
  console.log(`${indent}---`);
}

function walkTree(element, automation, depth = 0, maxDepth = 3) {
  if (depth > maxDepth) return;
  
  inspectElement(element, depth);
  
  const children = element.findAll(
    TreeScopes.Children,
    automation.createTrueCondition()
  );
  
  for (const child of children) {
    walkTree(child, automation, depth + 1, maxDepth);
  }
}

const automation = new Automation();
const root = automation.getRootElement();
walkTree(root, automation);
```

## 📂 Project Structure

```
node-winautomation/
├── index.js                 # Module entry point
├── index.d.ts               # TypeScript definitions
├── binding.gyp              # Native build configuration
├── package.json
│
├── wrappers/                # Core COM wrappers
│   ├── IUIAutomationWrapper.*
│   ├── IUIAutomationElementWrapper.*
│   ├── IUIAutomationTreeWalkerWrapper.*
│   └── ...
│
├── patterns/                # Automation pattern implementations
│   ├── IUIAutomationInvokePatternWrapper.*
│   ├── IUIAutomationValuePatternWrapper.*
│   ├── IUIAutomationTextPatternWrapper.*
│   └── ... (30+ patterns)
│
├── enumerations/            # UI Automation enums
│   ├── ControlTypeIdsWrapper.*
│   ├── PropertyIdsWrapper.*
│   └── ...
│
├── utilities/               # Event handlers & helpers
│   ├── AutomationEventHandler.*
│   ├── FocusChangedEventHandler.*
│   └── ...
│
├── examples/                # Usage examples
│   ├── calculator-example.js
│   └── notepad-example.js
│
└── tests/                   # Test suite
    └── *.test.js
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

This project is based on the excellent work by **Chris Nimmons** and **Jake Cyr** in the [@bright-fish/node-ui-automation](https://github.com/bright-fish/node-ui-automation) project.

## 📄 License

MIT License - See [LICENSE](LICENSE) file for details.

## 📚 Resources

- [Windows UI Automation Overview](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
- [UI Automation Control Patterns](https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpatternsoverview)
- [Accessibility Insights for Windows](https://accessibilityinsights.io/docs/windows/overview/)
- [Node.js N-API Documentation](https://nodejs.org/api/n-api.html)

---

<div align="center">
Made with ❤️ for Windows automation
</div>
