# Google Workspace Marketplace Readiness Review

## Scope reviewed
- Manifest/scopes and add-on metadata
- API/auth and external request patterns
- Spreadsheet performance and trigger behavior
- UI security and release hygiene

## High-priority improvements
1. **Reduce OAuth scopes to least privilege**
   - Remove `https://www.googleapis.com/auth/script.send_mail` unless email sending is reintroduced.
   - Keep only scopes that are actively used by addon flows to improve Marketplace security review outcomes.

2. **Stop keeping identity/spreadsheet context as global constants**
   - `Session.getActiveUser().getEmail()` and active spreadsheet values are captured globally in `Code.js`.
   - Resolve these values inside functions (or via lightweight getter helpers) to avoid stale context and edge cases during trigger execution.

3. **Harden all UrlFetch flows**
   - Introduce a shared request wrapper with:
     - response code checks
     - JSON parse guards
     - retries with exponential backoff for 429/5xx
     - consistent timeout and structured error objects
   - Current API functions duplicate fetch logic and assume JSON bodies on all responses.

4. **Avoid passing sensitive credentials/tokens beyond necessity**
   - Plaid access token handling should stay server-side whenever possible.
   - Review payloads to ensure only minimal references are transmitted/stored.

5. **Production manifest hygiene**
   - Re-check `urlFetchWhitelist` entries before release (remove non-production domains/endpoints not needed in production).
   - Confirm add-on branding fields, privacy policy, terms, and support links are fully aligned with Marketplace listing requirements.

## Medium-priority optimizations
6. **Batch Spreadsheet operations aggressively**
   - Continue moving to `getValues()`/`setValues()` patterns and reduce per-cell operations in heavy paths (transactions/investments sync).
   - Add write chunking and single-pass mapping where possible.

7. **Trigger lifecycle safety**
   - Ensure install/update routines are idempotent and do not create duplicate triggers.
   - Add centralized trigger reconciliation utility and logging with trigger IDs.

8. **Observability and diagnostics**
   - Replace ad hoc logs with structured logs (event name, spreadsheetId, user hash, duration, status).
   - Add correlation IDs for external API calls and sync batches.

9. **Frontend cleanup for Marketplace quality**
   - Remove `console.log` statements from HTML client scripts.
   - Add user-facing error states for network/auth failures across setup and plaid connect flows.

10. **Repository/docs release readiness**
    - Fix README encoding/content and provide:
      - setup/deployment steps
      - required Script Properties/Secrets
      - production checklist
      - rollback instructions

## Suggested implementation order
1. Scope/manifest cleanup and external domain audit
2. Centralized fetch wrapper + error/retry model
3. Global context refactor (`UserEmail`, spreadsheet globals)
4. Token/data minimization pass for Plaid and billing payloads
5. Spreadsheet throughput pass (profiling + batching)
6. Logging/diagnostics standardization
7. Documentation and release checklist

## Optional pre-submission checklist
- [ ] Verify add-on works with least scopes only
- [ ] Verify all external domains are HTTPS + justified
- [ ] Verify no secrets/tokens are logged
- [ ] Verify no client `console.log` remains in shipped HTML
- [ ] Validate behavior with a new spreadsheet and existing spreadsheet
- [ ] Validate re-install/update path does not duplicate triggers
