# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TheFinU is a Google Apps Script (GAS) add-on for Google Sheets that integrates with Plaid for financial account management. It provides transaction syncing, investment tracking, balance history, budget reporting (monthly/yearly, individual/joint), and net worth calculations. Users connect bank accounts via Plaid Link, and the add-on writes financial data into structured sheets.

## Development & Deployment

- **Language**: Google Apps Script (V8 runtime), plain JavaScript, HTML templates
- **Deployment tool**: `clasp` (Google Apps Script CLI)
- **Push code**: `clasp push` (uploads all .js, .html, .json files to the GAS project)
- **Pull code**: `clasp pull` (downloads from GAS project)
- **Open in browser**: `clasp open` (opens the Apps Script editor)
- **No build step, no tests, no linting** — code runs directly in GAS runtime
- **Config**: `.clasp.json` defines the script ID and file extensions

## Architecture

### Core Entry Points
- **Code.js** — Main file. Contains `onOpen()`, `onInstall()`, `showSidebar()`, setup wizard logic, trigger management, template installation, sync orchestration, and UI template rendering
- **AppApiEndpoint.js** — All backend REST API calls (`validateUserSession`, `getAppSettings`, `getAppPlaidAccountById`, etc.). All endpoints use `getAuthHeaders()` for OAuth-based authentication
- **PlaidAppAPI.js** — Direct Plaid API calls (`generatePlaidTokenLink`, `getPlaidTransactionSyncData`, `getPlaidAccountBalance`, `getPlaidInvestmentsData`, `plaidRequest`)

### Data Sheet Modules
Each module reads/writes to a specific sheet:
- **Transactions.js** — Transaction sync, insert, update, clear operations on the "Transactions" sheet
- **Investments.js** — Investment holdings on the "Investments" sheet
- **BalanceHistory.js** — Balance history tracking on the "Balance History" sheet
- **Accounts.js** — Account sheet data management
- **NetWorth.js / JointNetWorth.js** — Net worth report generation
- **MonthlyBudget.js / JointMonthlyBudget.js** — Monthly budget reports
- **YearlyBudget.js / JointYearlyBudget.js** — Yearly budget reports
- **Report.js** — Report generation orchestration

### HTML Templates (Sidebar UI)
- **Index.html / UserIndex.html** — Setup wizard and main dashboard sidebar containers
- **SetupWizard.html** — Step-by-step onboarding wizard
- **UserDashboard.html** — Main user interface after setup
- **AccountListCard.html / AccountDetailCard.html** — Account management views
- **ConnectPlaidAccount.html / LinkImportRunner.html** — Plaid Link flow and data import
- **setupWizardJs.html / userAppJs.html / commonJs.html** — Client-side JavaScript (included via `<?!= include() ?>`)

### Key Patterns
- **Global variables**: `UserEmail`, `UserSpreadsheet` are lazily initialized at the top of Code.js with try/catch for trigger safety
- **Sheet name constants**: `USER_TRANSACTIONS_SHEET`, `USER_BALANCE_HISTORY_SHEET`, etc. defined at top of Code.js
- **Template rendering**: Server-side functions return HTML strings via `HtmlService.createTemplateFromFile().evaluate().getContent()`, injected into sidebar via `innerHTML`
- **GAS templating**: `<? ?>` for scriptlets, `<?= ?>` for HTML-escaped output, `<?!= ?>` for raw output. NEVER use `<% %>` or `<%= %>` (EJS/ERB syntax) — GAS only recognizes `<? ?>` variants. Use `<?= ?>` for user-controlled data; use data attributes + JS event listeners instead of inline `onclick` with dynamic values
- **Backend API**: All calls go to `API_ENDPOINT` (`https://thefinu.stallioni.com/`). Responses follow `{ success: boolean, result: ... }` pattern
- **Settings caching**: `getAppSettings()` is cached per script execution via `_cachedAppSettings`. Call `clearAppSettingsCache()` if you need fresh data
- **Auto-sync trigger**: `runThefinUPlaidAutoSync()` runs daily via time-based trigger. The installable edit trigger uses `handleAddonEdit()`
- **Stripe checkout**: Handled server-side via backend endpoint `api/payment/create-checkout-session`

## Important Conventions

- All sheet writes use Comfortaa font, 9pt, black, bold formatting
- Budget reports use the "Definition" sheet for column mappings and configuration
- The "Categories" sheet drives budget category structure (type → group → category hierarchy)
- Joint reports split amounts between two people using allocation percentages from the Categories sheet
- `reApplyFormulaToSpreadsheet(sheetName)` resets formulas after template installation — each case handles a specific sheet

## Backend API Requirements

The backend at `thefinu.stallioni.com` must validate the `Authorization: Bearer <OAuthToken>` and `X-User-Email` headers sent by `getAuthHeaders()`. The backend must also implement `POST api/payment/create-checkout-session` to handle Stripe session creation server-side (never expose Stripe secret keys to client code).
