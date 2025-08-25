import { test, expect } from '@playwright/test';
import * as fs from 'fs';

const creds = JSON.parse(fs.readFileSync('./config/credentials.json', 'utf-8'));

test('Login to Office.com', async ({ page }) => {
  // 1. Going on the site
  await page.goto('https://www.office.com/');

  // 2. Click on Sign in
  await page.click('text=Sign in');

  // 3. Insert email
  await page.fill('input[type="email"]', creds.email);
  await page.click('input[type="submit"]');

  // 4. Insert password
  await page.fill('input[type="password"]', creds.password);
  await page.click('input[type="submit"]');

  // 5. In some cases it appears "Stay signed in?"
  try {
    await page.click('text=Yes', { timeout: 5000 });
  } catch {}

  // 6. Check that we arrived logged (that "Excel" is present)
  await expect(page.locator('body')).toContainText('Excel');
});
