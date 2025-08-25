# excel-playwright-test
## Excel Online TODAY() End-to-End Test Automation

---

## Overview

This project explores the automation of writing formulas into an Excel Online instance (opened inside a browser tab/iframe). The main idea is to automatically select a cell, input a formula (e.g., `=TODAY()`), and validate whether Excel evaluates it correctly.

---

## Current Status

- The script successfully opens Excel Online and navigates to a cell.
- A formula is typed into the formula bar, but after pressing Enter, the content disappears or is not applied to the selected cell.
- Manual typing works fine; the issue only occurs in automation.
- Direct cell selection inside the Excel grid was not realized due to Excel Online rendering via canvas and shadow DOM.
- The test inputs the formula in the formula bar instead.

---

## Challenges

- **Focus management:** Excel’s formula bar does not behave like a standard input field; it requires correct focus handling.
- **Iframe isolation:** Excel is rendered in an iframe/grid structure, making selectors unstable.

---

## Assumptions

It seems possible to achieve the goal by:

1. Mapping precise cell coordinates in the grid, OR
2. Using Microsoft Graph API for direct formula insertion.

---

## Project Structure

- `credentials/` – Configuration data for login (username, password).  
- `storageState.json` – Stores browser login session for automatic authentication.  
- `login-only.spec.ts` – Test to validate login process.  
- `today-excel.spec.ts` – Main test for the `TODAY()` function.  
- `package.json`, `package-lock.json` – Node project files.  
- `playwright.config.ts` – Playwright configuration file.  

---

## Demo

A recorded video demo of the test execution is available in the `demo/` folder.

---

# How to Run

# 1. Clone the repository
git clone <repository-url>
cd excel-playwright-test

# 2. Install dependencies
npm install

# 3. Run the test
npx playwright test

# Optional: run in headed mode by setting headless: false in the test file

# FAQ

Q1: Why doesn’t the test write directly into the cell?

Excel Online uses a canvas for the grid and a shadow DOM, which prevents reliable clicking and typing in arbitrary cells. The formula bar is used as a workaround.

Q2: Can this test be extended for other formulas?

Yes, by changing the formula in today-excel.spec.ts and updating the expected result parsing.

Q3: Do I need to login manually?

No, the test uses storageState.json to reuse a saved authenticated session.

Q4: Why are there delays in the test?

Delays ensure the workbook fully loads and the formula bar is ready to accept input before typing.

Q5: Why does the formula disappear after pressing Enter?

Because Excel Online’s formula bar does not act like a regular HTML <input>. It requires specific event triggers (blur, commit, grid focus), which Playwright does not automatically simulate.

Q6: Could this work with exact cell coordinates?

Yes, in theory. If the iframe and grid system were fully mapped, automation could click directly on a cell at (x, y) coordinates and input the formula.

Q7: Why not use Graph API?

Graph API is the recommended Microsoft way to read/write Excel cells programmatically. It requires authentication via Azure AD, proper access permissions, and replacing UI automation with API calls. The project scope was limited to UI automation.

Q8: Is this a bug in the automation tool?

No, it is a limitation of Excel Online’s rendering. Playwright works correctly but cannot fully simulate Excel’s internal event handling.

Q9: What’s next?

Investigate direct Graph API integration.

Try cell coordinate mapping using screen coordinates instead of selectors.

Explore custom keyboard event simulation to replicate manual input more closely.

Notes
All test results and outputs are logged to the console.
The repository is ready for submission with all required files included.
