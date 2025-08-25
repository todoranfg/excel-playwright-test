import { chromium } from '@playwright/test';
import readline from 'readline';

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto('https://www.office.com/');

  console.log('Login manually in your Office account and then write "ok" and press Enter in  the terminal...');

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  rl.question('', async (answer) => {
    if (answer.toLowerCase() === 'ok') {
      await context.storageState({ path: 'storageState.json' });
      console.log('Login was saved!');
      await browser.close();
      rl.close();
      process.exit();
    } else {
      console.log('You need to write "ok" after you login manually.');
      rl.close();
      process.exit(1);
    }
  });
})();

