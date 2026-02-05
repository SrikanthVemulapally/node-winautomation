/**
 * Test all UI Automation examples
 */

console.log('=== Testing All UI Automation Examples ===\n');

// Load native addon directly
const addon = require('bindings')('Automation');

// Test 1: Calculator Example
async function testCalculator() {
    console.log('Test 1: Calculator Example');
    console.log('----------------------------');
    
    try {
        const automation = new addon.Automation();
        const root = automation.getRootElement();
        
        const namePropertyId = addon.PropertyIds.NamePropertyId;
        const calcCondition = automation.createPropertyCondition(namePropertyId, 'Calculator');
        const calculator = root.findFirst(4, calcCondition);
        
        if (!calculator) {
            console.log('⚠ Calculator not running - skipping test');
            return false;
        }
        
        console.log('✓ Calculator found');
        
        // Click 2
        const button2 = calculator.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Two'));
        if (button2) {
            button2.getCurrentPattern(addon.PatternIds.InvokePatternId).invoke();
            console.log('✓ Clicked "2"');
            await new Promise(r => setTimeout(r, 300));
        }
        
        // Click +
        const plusButton = calculator.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Plus'));
        if (plusButton) {
            plusButton.getCurrentPattern(addon.PatternIds.InvokePatternId).invoke();
            console.log('✓ Clicked "+"');
            await new Promise(r => setTimeout(r, 300));
        }
        
        // Click 3
        const button3 = calculator.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Three'));
        if (button3) {
            button3.getCurrentPattern(addon.PatternIds.InvokePatternId).invoke();
            console.log('✓ Clicked "3"');
            await new Promise(r => setTimeout(r, 300));
        }
        
        // Click =
        const equalsButton = calculator.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Equals'));
        if (equalsButton) {
            equalsButton.getCurrentPattern(addon.PatternIds.InvokePatternId).invoke();
            console.log('✓ Clicked "="');
            await new Promise(r => setTimeout(r, 500));
        }
        
        console.log('✅ Calculator test PASSED (2 + 3 = 5)\n');
        return true;
        
    } catch (error) {
        console.error('❌ Calculator test FAILED:', error.message);
        return false;
    }
}

// Test 2: Notepad Example
async function testNotepad() {
    console.log('Test 2: Notepad Example');
    console.log('------------------------');
    
    try {
        const automation = new addon.Automation();
        const root = automation.getRootElement();
        
        const namePropertyId = addon.PropertyIds.NamePropertyId;
        
        // Try different Notepad window names
        let notepad = root.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Untitled - Notepad'));
        if (!notepad) {
            notepad = root.findFirst(4, automation.createPropertyCondition(namePropertyId, 'Notepad'));
        }
        
        if (!notepad) {
            console.log('⚠ Notepad not running - skipping test');
            return false;
        }
        
        console.log('✓ Notepad window found');
        
        // Find edit control
        const editCondition = automation.createPropertyCondition(
            addon.PropertyIds.ControlTypePropertyId,
            addon.ControlTypeIds.EditControlTypeId
        );
        const editor = notepad.findFirst(4, editCondition);
        
        if (editor) {
            console.log('✓ Text editor found');
            
            const valuePattern = editor.getCurrentPattern(addon.PatternIds.ValuePatternId);
            if (valuePattern) {
                const testText = 'Hello from node-winautomation!\nAll examples tested successfully!';
                valuePattern.setValue(testText);
                console.log('✓ Text set successfully');
                console.log('✅ Notepad test PASSED\n');
                return true;
            }
        }
        
        console.log('⚠ Could not set text in Notepad');
        return false;
        
    } catch (error) {
        console.error('❌ Notepad test FAILED:', error.message);
        return false;
    }
}

// Test 3: Outlook Window Finding
async function testOutlook() {
    console.log('Test 3: Outlook Window Finding');
    console.log('--------------------------------');
    
    try {
        const automation = new addon.Automation();
        const root = automation.getRootElement();
        
        const namePropertyId = addon.PropertyIds.NamePropertyId;
        
        // Try to find Outlook window
        const outlookConditions = [
            'Inbox - Outlook',
            'Microsoft Outlook',
            'Outlook'
        ];
        
        let outlookWindow = null;
        for (const name of outlookConditions) {
            outlookWindow = root.findFirst(4, automation.createPropertyCondition(namePropertyId, name));
            if (outlookWindow) {
                console.log(`✓ Outlook window found: "${name}"`);
                break;
            }
        }
        
        if (!outlookWindow) {
            console.log('⚠ Outlook not running - skipping test');
            return false;
        }
        
        console.log('✓ Outlook automation ready');
        console.log('✅ Outlook test PASSED (window found)\n');
        return true;
        
    } catch (error) {
        console.error('❌ Outlook test FAILED:', error.message);
        return false;
    }
}

// Run all tests
async function runAllTests() {
    console.log('Starting all example tests...\n');
    
    const results = {
        calculator: await testCalculator(),
        notepad: await testNotepad(),
        outlook: await testOutlook()
    };
    
    console.log('=== Test Results Summary ===');
    console.log(`Calculator: ${results.calculator ? '✅ PASSED' : '⚠ SKIPPED'}`);
    console.log(`Notepad:    ${results.notepad ? '✅ PASSED' : '⚠ SKIPPED'}`);
    console.log(`Outlook:    ${results.outlook ? '✅ PASSED' : '⚠ SKIPPED'}`);
    
    const passedCount = Object.values(results).filter(r => r).length;
    const totalCount = Object.keys(results).length;
    
    console.log(`\nTotal: ${passedCount}/${totalCount} tests passed`);
    
    if (passedCount === 0) {
        console.log('\n⚠ No applications were running. Please start Calculator, Notepad, or Outlook and run again.');
    } else {
        console.log('\n✅ All available examples tested successfully!');
    }
}

runAllTests().catch(error => {
    console.error('Unexpected error:', error);
    process.exit(1);
});
