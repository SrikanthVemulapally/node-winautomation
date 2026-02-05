/**
 * Fixed Notepad Test - Properly finding and setting text
 */

console.log('=== Notepad Automation Test (Fixed) ===\n');

const addon = require('bindings')('Automation');

async function testNotepadFixed() {
    try {
        console.log('1. Creating Automation instance...');
        const automation = new addon.Automation();
        const root = automation.getRootElement();
        
        console.log('\n2. Finding Notepad window...');
        const namePropertyId = addon.PropertyIds.NamePropertyId;
        
        // Try different window names
        let notepad = root.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Untitled - Notepad'));
        if (!notepad) {
            notepad = root.findFirst(4, automation.createPropertyCondition(namePropertyId, '*Notepad'));
        }
        if (!notepad) {
            // Try finding any window with Notepad in the name
            const windows = root.findAll(4, automation.createPropertyCondition(
                addon.PropertyIds.ControlTypePropertyId,
                addon.ControlTypeIds.WindowControlTypeId
            ));
            
            for (let i = 0; i < windows.length; i++) {
                const win = windows[i];
                try {
                    const name = win.currentName;
                    if (name && name.toLowerCase().includes('notepad')) {
                        notepad = win;
                        console.log(`✓ Found window: "${name}"`);
                        break;
                    }
                } catch (e) {
                    // Skip windows we can't access
                }
            }
        }
        
        if (!notepad) {
            console.log('✗ Notepad not found. Please start Notepad.');
            return;
        }
        
        console.log('✓ Notepad window found');
        
        console.log('\n3. Finding text editor control...');
        
        // Method 1: Try Document control type
        let editor = notepad.findFirst(4, automation.createPropertyCondition(
            addon.PropertyIds.ControlTypePropertyId,
            addon.ControlTypeIds.DocumentControlTypeId
        ));
        
        if (editor) {
            console.log('✓ Found Document control');
        } else {
            // Method 2: Try Edit control type
            console.log('  Trying Edit control...');
            editor = notepad.findFirst(4, automation.createPropertyCondition(
                addon.PropertyIds.ControlTypePropertyId,
                addon.ControlTypeIds.EditControlTypeId
            ));
            
            if (editor) {
                console.log('✓ Found Edit control');
            }
        }
        
        if (!editor) {
            // Method 3: Find all edit controls and use the largest one
            console.log('  Searching all edit controls...');
            const allEdits = notepad.findAll(4, automation.createPropertyCondition(
                addon.PropertyIds.ControlTypePropertyId,
                addon.ControlTypeIds.EditControlTypeId
            ));
            
            console.log(`  Found ${allEdits.length} edit controls`);
            
            if (allEdits.length > 0) {
                // Use the first edit control (usually the main text area)
                editor = allEdits[0];
                console.log('✓ Using first edit control');
            }
        }
        
        if (!editor) {
            console.log('✗ Could not find text editor control');
            return;
        }
        
        console.log('\n4. Setting text...');
        
        // Try Value pattern first
        try {
            const valuePattern = editor.getCurrentPattern(addon.PatternIds.ValuePatternId);
            if (valuePattern) {
                const testText = 'Hello from node-winautomation!\n\nNotepad automation is working!\n\nTest completed at: ' + new Date().toLocaleString();
                valuePattern.setValue(testText);
                console.log('✓ Text set using Value pattern');
                console.log('✅ Notepad automation successful!\n');
                return;
            }
        } catch (e) {
            console.log('  Value pattern not available:', e.message);
        }
        
        // Try Text pattern
        try {
            const textPattern = editor.getCurrentPattern(addon.PatternIds.TextPatternId);
            if (textPattern) {
                console.log('✓ Text pattern available (read-only)');
                console.log('  Note: Text pattern is typically read-only');
            }
        } catch (e) {
            console.log('  Text pattern not available');
        }
        
        // Try setting focus and using SendKeys (alternative method)
        console.log('\n5. Trying alternative method (focus + keyboard)...');
        try {
            editor.setFocus();
            console.log('✓ Focus set on editor');
            console.log('⚠ For keyboard input, you would need SendKeys or similar');
            console.log('  The control is ready for keyboard input');
        } catch (e) {
            console.log('  Could not set focus:', e.message);
        }
        
        console.log('\n✅ Notepad window and editor found successfully!');
        console.log('   Note: Value pattern may not be available on all Notepad versions.');
        
    } catch (error) {
        console.error('\n❌ Test failed:', error.message);
        console.error(error.stack);
    }
}

testNotepadFixed().then(() => {
    console.log('\n=== Test Complete ===');
});
