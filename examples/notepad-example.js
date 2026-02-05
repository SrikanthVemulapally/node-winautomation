/**
 * Notepad Automation Example
 *
 * This example demonstrates how to:
 * 1. Find a Notepad window
 * 2. Locate the text editor control
 * 3. Set text using the Value pattern
 *
 * Prerequisites:
 * - Have Notepad running (notepad.exe)
 * - Build the addon first: npm run build
 */

const { Automation, PropertyIds, ControlTypeIds, TreeScopes, PatternIds } = require('..');

async function automateNotepad() {
  try {
    console.log('Initializing UI Automation...');
    const automation = new Automation();
    const root = automation.getRootElement();

    // Find Notepad window by name
    console.log('Looking for Notepad window...');
    const windowCondition = automation.createPropertyCondition(
      PropertyIds.NamePropertyId,
      'Untitled - Notepad'
    );

    const notepadWindow = root.findFirst(
      TreeScopes.Children,
      windowCondition
    );

    if (!notepadWindow) {
      console.log('Notepad window not found. Please open Notepad and try again.');
      return;
    }

    console.log('Found Notepad window:', notepadWindow.currentName);

    // Find the edit control (text area)
    // Windows 11 Notepad uses Document control type, older versions use Edit
    console.log('Looking for edit control...');
    
    // Try Document first (Windows 11), then Edit (older Windows)
    let editControl = notepadWindow.findFirst(
      TreeScopes.Descendants,
      automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.DocumentControlTypeId)
    );
    
    if (!editControl) {
      editControl = notepadWindow.findFirst(
        TreeScopes.Descendants,
        automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.EditControlTypeId)
      );
    }

    if (!editControl) {
      console.log('Edit control not found in Notepad window.');
      return;
    }

    console.log('Found edit control:', editControl.currentName || 'Text Editor');

    // Get the Value pattern to set text
    console.log('Getting Value pattern...');
    const valuePattern = editControl.getCurrentPattern(
      PatternIds.ValuePatternId
    );

    if (!valuePattern) {
      console.log('Value pattern not available for this control.');
      return;
    }

    // Set text in Notepad
    const message = `Hello from node-winautomation!

This text was automatically inserted using the Windows UI Automation API.

Timestamp: ${new Date().toISOString()}

The automation can:
- Find windows and controls
- Set and get text values
- Click buttons
- Handle various UI patterns
- And much more!

Visit the README.md for more examples.
`;

    console.log('Setting text in Notepad...');
    valuePattern.setValue(message);

    console.log('✓ Successfully automated Notepad!');
    console.log('Check Notepad window for the inserted text.');

  } catch (error) {
    console.error('Error during automation:', error.message);
    console.error(error);
  }
}

// Run the automation
console.log('=== Notepad Automation Example ===\n');
automateNotepad().then(() => {
  console.log('\nExample completed.');
}).catch(err => {
  console.error('Unexpected error:', err);
  process.exit(1);
});
