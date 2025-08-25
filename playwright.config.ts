import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  use: {
    browserName: 'chromium', // run only on Chromium (Chrome)
    headless: false,         // watch the browser
    screenshot: 'only-on-failure',
    video: 'retain-on-failure'
  },
});
