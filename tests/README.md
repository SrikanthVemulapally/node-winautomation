# Tests

## Unit Tests (Jest)
- `Enumerations.test.js` - Tests for UI Automation enumerations
- `IUIAutomation.test.js` - Tests for Automation class
- `IUIAutomationElement.test.js` - Tests for AutomationElement
- `IUIAutomationTreeWalker.test.js` - Tests for TreeWalker
- `Shared.js` - Shared test utilities
- `jest.config.js` - Jest configuration

## Integration Tests
- `test-all-examples.js` - Tests all UI Automation examples (Calculator, Notepad, Outlook)
- `test-build.js` - Verifies build and module loading
- `test-features.js` - Tests UI Automation and COM Automation features
- `test-notepad-fixed.js` - Notepad automation test

## Running Tests

### Unit Tests
```bash
cd tests
npm test
```

### Integration Tests
```bash
# Test all examples
node tests/test-all-examples.js

# Test build
node tests/test-build.js

# Test features
node tests/test-features.js
```
