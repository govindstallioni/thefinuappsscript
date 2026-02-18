/**
 * Retrieves user data and validates it against your external API.
 * Called from the sidebar when it loads.
 */
function validateUserSession() {
  try {
    const userEmail = getUserEmail();
    const spreadsheetId = getUserSpreadsheetId();

    if (!userEmail) {
      throw new Error('Email access denied. Please re-authorize the add-on.');
    }

    const apiUrl = API_ENDPOINT + 'api/users/validate-user';
    const payload = {
      email: userEmail,
      spreadsheetId: spreadsheetId,
      timestamp: new Date().toISOString(),
    };

    const response = requestJson(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    });

    const result = response.body || {};

    if (result && result.data) {
      const d = result.data;
      const found = d.subscriptionId || d.subscription_id || (d.subscriptions && d.subscriptions.id) || (d.subscription && d.subscription.id) || d.subId || null;
      if (found) {
        d.subscriptionId = found;
      }
    }

    return {
      success: response.success,
      result: result,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    Logger.log(JSON.stringify(e, null, 2));
    return { success: false, error: e.toString() };
  }
}

function getAppSettings() {
  try {
    const apiUrl = API_ENDPOINT + 'api/settings';
    const response = requestJson(apiUrl, { method: 'get' });
    return {
      success: response.success,
      result: response.body,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
    };
  }
}

function storePlaidAPIAccounts(data) {
  try {
    const userEmail = getUserEmail();

    if (!userEmail) {
      throw new Error('Email access denied. Please re-authorize the add-on.');
    }

    const apiUrl = API_ENDPOINT + 'api/accounts/store-plaid';
    const response = requestJson(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(data),
    });

    return {
      success: response.success,
      result: response.body,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAppPlaidConnectedAccounts() {
  try {
    const apiUrl = API_ENDPOINT + 'api/accounts/get-by-email/' + encodeURIComponent(getUserEmail());
    const response = requestJson(apiUrl, { method: 'get' });

    return {
      success: response.success,
      result: response.body,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
    };
  }
}

function getAppPlaidAccountById(accountId) {
  try {
    const apiUrl = API_ENDPOINT + 'api/accounts/' + accountId;
    const response = requestJson(apiUrl, { method: 'get' });

    return {
      success: response.success,
      result: response.body,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
    };
  }
}

function updateAppAccountDetailById(accountId, data) {
  try {
    const apiUrl = API_ENDPOINT + 'api/accounts/update-account/' + accountId;
    const response = requestJson(apiUrl, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(data),
    });

    return {
      success: response.success,
      result: response.body,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
    };
  }
}

/**
 * Cancels the user's subscription by notifying the external API.
 */
function confirmCancelUserSubscription() {
  try {
    const apiUrl = API_ENDPOINT + 'api/payment/unsubscribe';
    const payload = {
      email: getUserEmail(),
    };

    const response = requestJson(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    });

    if (response.success) {
      try { clearSubscriptionProgress(); } catch (e) {}
    }

    return {
      success: response.success,
      result: response.body,
      statusCode: response.statusCode,
      error: response.error || null,
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
