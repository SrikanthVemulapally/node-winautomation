/**
 * Calculator Automation Example
 *
 * This example demonstrates how to automate the Windows Calculator app:
 * 1. Find the Calculator window
 * 2. Click number buttons
 * 3. Perform a calculation (2 + 3 = 5)
 *
 * Prerequisites:
 * - Windows 10/11 Calculator app running (calc.exe or Calculator from Start menu)
 * - Build the addon first: npm run build
 */

const { UIAutomation } = require('../../../index.js');

async function clickButton(window, buttonName, automation) {
  console.log(`Looking for button: ${buttonName}`);

  const buttonCondition = automation.createPropertyCondition(
    UIAutomation.PropertyIds.NamePropertyId,
    buttonName
  );

  const button = window.findFirst(
    UIAutomation.TreeScopes.Descendants,
    buttonCondition
  );

  if (!button) {
    console.log(`Button "${buttonName}" not found`);
    return false;
  }

  console.log(`Clicking button: ${buttonName}`);
  const invokePattern = button.getCurrentPattern(UIAutomation.PatternIds.InvokePatternId);

  if (!invokePattern) {
    console.log(`Invoke pattern not available for button: ${buttonName}`);
    return false;
  }

  invokePattern.invoke();

  // Small delay to allow UI to update
  await new Promise(resolve => setTimeout(resolve, 200));

  return true;
}

async function automateCalculator() {
  try {
    console.log('Initializing UI Automation...');
    const automation = new UIAutomation.Automation();
    const root = automation.getRootElement();

    // Find Calculator window
    console.log('Looking for Calculator window...');
    const calcCondition = automation.createPropertyCondition(
      UIAutomation.PropertyIds.NamePropertyId,
      'Calculator'
    );

    const calcWindow = root.findFirst(
      UIAutomation.TreeScopes.Children,
      calcCondition
    );

    if (!calcWindow) {
      console.log('Calculator window not found. Please open Calculator and try again.');
      console.log('You can open it by searching "Calculator" in the Start menu.');
      return;
    }

    console.log('Found Calculator window:', calcWindow.currentName);
    console.log('\nPerforming calculation: 2 + 3 = 5\n');

    // Perform calculation: 2 + 3 = 5
    await clickButton(calcWindow, 'Two', automation);
    await clickButton(calcWindow, 'Plus', automation);
    await clickButton(calcWindow, 'Three', automation);
    await clickButton(calcWindow, 'Equals', automation);

    console.log('\n✓ Successfully automated Calculator!');
    console.log('The result (5) should now be displayed in Calculator.');

  } catch (error) {
    console.error('Error during automation:', error.message);
    console.error(error);
  }
}

// Run the automation
console.log('=== Calculator Automation Example ===\n');
automateCalculator().then(() => {
  console.log('\nExample completed.');
}).catch(err => {
  console.error('Unexpected error:', err);
  process.exit(1);
});
