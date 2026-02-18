function runMarketplacePreflightChecks() {
  const results = {
    passed: [],
    warnings: [],
    failed: [],
    timestamp: new Date().toISOString(),
  };

  // 1) Active spreadsheet context
  try {
    const ss = getUserSpreadsheet();
    if (ss && ss.getId()) {
      results.passed.push('Active spreadsheet context available.');
    } else {
      results.failed.push('Active spreadsheet context is not available.');
    }
  } catch (e) {
    results.failed.push('Unable to access active spreadsheet: ' + e.toString());
  }

  // 2) User email availability
  try {
    const email = getUserEmail();
    if (email) {
      results.passed.push('User email scope and context available.');
    } else {
      results.warnings.push('User email is empty; verify OAuth grant and domain sharing settings.');
    }
  } catch (e) {
    results.warnings.push('Unable to read user email: ' + e.toString());
  }

  // 3) Required sheets present
  try {
    const ss = getUserSpreadsheet();
    const allSheets = ss.getSheets().map(s => s.getName());
    const required = appBaseTemplates();
    const missing = required.filter(name => allSheets.indexOf(name) === -1);

    if (missing.length === 0) {
      results.passed.push('All required base sheets exist.');
    } else {
      results.failed.push('Missing base sheets: ' + missing.join(', '));
    }
  } catch (e) {
    results.failed.push('Sheet inventory check failed: ' + e.toString());
  }

  // 4) Trigger duplication safety check
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const byHandler = {};
    triggers.forEach(function (t) {
      const h = t.getHandlerFunction();
      byHandler[h] = (byHandler[h] || 0) + 1;
    });

    const duplicates = Object.keys(byHandler).filter(function (h) { return byHandler[h] > 1; });

    if (duplicates.length === 0) {
      results.passed.push('No duplicate project triggers detected.');
    } else {
      results.warnings.push('Duplicate triggers detected for: ' + duplicates.join(', '));
    }
  } catch (e) {
    results.warnings.push('Trigger duplication check failed: ' + e.toString());
  }

  // 5) Backend settings availability
  try {
    const settings = getAppSettings();
    if (settings && settings.success) {
      results.passed.push('Backend settings endpoint reachable.');
    } else {
      results.failed.push('Backend settings endpoint failed.');
    }
  } catch (e) {
    results.failed.push('Backend settings check failed: ' + e.toString());
  }

  return results;
}

function runMarketplacePreflightChecksUI() {
  const ui = SpreadsheetApp.getUi();
  const report = runMarketplacePreflightChecks();

  const lines = [];
  lines.push('Preflight completed at: ' + report.timestamp);
  lines.push('Passed: ' + report.passed.length);
  lines.push('Warnings: ' + report.warnings.length);
  lines.push('Failed: ' + report.failed.length);

  if (report.failed.length > 0) {
    lines.push('');
    lines.push('Failures:');
    report.failed.forEach(function (f) { lines.push('- ' + f); });
  }

  if (report.warnings.length > 0) {
    lines.push('');
    lines.push('Warnings:');
    report.warnings.forEach(function (w) { lines.push('- ' + w); });
  }

  ui.alert('Marketplace Preflight Report', lines.join('\n'), ui.ButtonSet.OK);
  return report;
}
