# Google Workspace Marketplace Submission Checklist

## 1) Technical preflight (in Sheets Add-on)
- Run **Add-ons → TheFinU Addon → Run Marketplace Preflight**.
- Confirm:
  - No failures.
  - No trigger duplication warnings.
  - Backend settings endpoint is reachable.

## 2) OAuth and scopes
- Confirm only required scopes are in `appsscript.json`.
- Re-authorize and test critical flows after scope changes.

## 3) Functional validation
- New user flow:
  - Setup wizard
  - Subscription check
  - Template setup
  - Plaid link and import
- Existing user flow:
  - Open dashboard
  - Sync transactions/investments
  - Unlink and relink account

## 4) Reliability checks
- Validate retry behavior for temporary backend errors.
- Confirm user-facing errors are shown for failed API requests.
- Confirm no duplicate triggers after reinstall/update.

## 5) Release metadata (Marketplace)
- App name, logo, and descriptions finalized.
- Privacy policy URL and Terms URL finalized.
- Support contact email and website finalized.
- Test account credentials prepared for Google reviewers.

## 6) Final go/no-go
- `runMarketplacePreflightChecks()` returns zero failures.
- Core journeys pass end-to-end on a clean spreadsheet.
- Release notes documented.
