import { test, chromium } from '@playwright/test';
import fs from 'fs';

test('Excel TODAY() via formula bar', async ({}, testInfo) => {
  testInfo.setTimeout(30000);

  console.log('Launching Chromium browser...');
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: 'storageState.json' });
  const page = await context.newPage();

  console.log('Navigating to Excel Online...');
  await page.goto('https://www.office.com/launch/excel');
  await page.waitForLoadState('domcontentloaded');

  console.log('Opening Blank Workbook...');
  const [excelPage] = await Promise.all([
    context.waitForEvent('page'),
    page.locator('text=Blank workbook').first().click(),
  ]);

  await excelPage.waitForLoadState('domcontentloaded');
  await excelPage.waitForTimeout(10000); // Wait for workbook to fully initialize

  console.log('Selecting main Excel iframe...');
  const iframe = await excelPage.waitForSelector('iframe[name^="WacFrame_Excel"]', { state: 'attached', timeout: 10_000 });
  if (!iframe) {
    console.log('❌ Error: Excel iframe not found or not loaded!');
    await browser.close();
    return;
  }

  const excelFrame = await iframe.contentFrame();
  if (!excelFrame) {
    console.log('❌ Error: Could not access Excel iframe content!');
    await browser.close();
    return;
  }

  console.log('Focusing on the formula bar and typing =TODAY()...');
  let formulaBar = excelFrame.locator('[aria-label="Formula Bar"] input').first();
  let isVisible = await formulaBar.isVisible({ timeout: 5000 }).catch(() => false);

  if (!isVisible) {
    console.log('❌ Original locator did not work, trying alternative locator...');
    formulaBar = excelFrame.locator('[role="textbox"]').first();
    isVisible = await formulaBar.isVisible({ timeout: 5000 }).catch(() => false);
    if (!isVisible) {
      console.log('❌ Alternative locator failed. Inspecting DOM...');
      const allElements = await excelFrame.locator('*').all();
      console.log('Number of elements found in iframe:', allElements.length);
      await browser.close();
      return;
    } else {
      console.log('✅ Found formula bar using [role="textbox"].');
    }
  } else {
    console.log('✅ Found formula bar using aria-label.');
  }

  // Focus and enter the formula
  await formulaBar.click();
  await formulaBar.fill(''); // Clear existing content
  await formulaBar.fill('=TEXT(TODAY(),"yyyy-mm-dd")');
  await formulaBar.press('Enter');

  await excelPage.waitForTimeout(4000); // Wait for value to update

  console.log('Reading the value of the active cell...');
  const activeCell = excelFrame.locator('div[aria-selected="true"]');
  const value = await activeCell.innerText();
  console.log('Active cell value:', value);

  const today = new Date().toLocaleDateString('en-CA'); // YYYY-MM-DD
  if (value.includes(today)) {
    console.log('✅ Test passed: TODAY() is correct.');
  } else {
    console.log('❌ Test failed: TODAY() does not match.');
    console.log('Cell value:', value, 'Expected:', today);
  }

  await browser.close();
});

