async function generatePlaidTokenLink( access_token = null ) {
  
  let response = [];
  const appSettingsData = getAppSettings();

  if( appSettingsData.success === true ){

    const plaidEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/link/token/create';

    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      user: getAppUserId(),
      client_name: 'ThefinU, LLC',
      products: ['transactions'], // The Plaid products you want to access
      optional_products: ['investments'],
      transactions: {
        days_requested: 730
      },
      webhook: appSettingsData.result.plaidWebhookUrl,
      country_codes: ['US'], // Country codes for available institutions
      language: 'en', // Language for the Link interface
      access_token : access_token,
      update: { account_selection_enabled: true }
    };

    response = plaidRequest(plaidEndpoint, payload);
  }
  return response;
}

/**
 * Exchange public_token → access_token + item_id + accounts
 */
function exchangePublicTokenForAccessToken(public_token, metadata) {

  const appSettingsData = getAppSettings();

  if( appSettingsData.success === true ){

    const plaidExchangePublicTokenEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/item/public_token/exchange';

    const payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      public_token: public_token
    };

    let response = plaidRequest(plaidExchangePublicTokenEndpoint, payload);

    const access_token = response.access_token;
    const item_id = response.item_id;

    // Prepare payload for your backend
    const dataToSend = {
      timestamp: new Date().toISOString(),
      email: Session.getActiveUser().getEmail(),
      plaid_item_id: item_id,
      access_token: access_token,           // ← usually you keep this in your DB, not send!
      accounts: metadata.accounts,
      metadata: {
        institution_name: metadata.institution.name,
        institution_id: metadata.institution.institution_id,
        link_session_id: metadata.link_session_id,
        accounts_count: metadata.accounts?.length || 0
      }
    };

    updatePlaidAccountIDOnSheets( metadata.institution.institution_id, metadata.accounts );
    // Send to YOUR backend
    storePlaidAPIAccounts(dataToSend);
    // For demo: we can return something useful to UI
    return {
      success: true,
      item_id: item_id,
      message: 'New Accounts Added Successfully'
    };

  }else{
    return {
      success: false,
      message: 'Someting went wrong, please try again.'
    };
  }
}

function getPlaidTransactionSyncData( account_id, new_cursor = null ){

  let response = { error: true, error_message: 'Account or settings lookup failed' };
  const accountData = getAppPlaidAccountById(account_id);
  const appSettingsData = getAppSettings();
  if( accountData.success === true && appSettingsData.success === true ){
    const plaidTransactionsEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/transactions/sync';
    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      access_token: accountData.result.access_token,
      cursor: new_cursor ? new_cursor : '',
      count: 500
    };
    response = plaidRequest(plaidTransactionsEndpoint, payload);
    // /transactions/sync returns all accounts for the item — filter to requested account
    if(response && !response.error){
      if(response.added) response.added = response.added.filter(function(t){ return t.account_id === account_id; });
      if(response.modified) response.modified = response.modified.filter(function(t){ return t.account_id === account_id; });
      if(response.removed) response.removed = response.removed.filter(function(t){ return t.account_id === account_id; });
    }
  }
  return response;
}

function getPlaidAccountBalance( account_id ){

  let response = { error: true, error_message: 'Account or settings lookup failed' };
  const accountData = getAppPlaidAccountById(account_id);
  const appSettingsData = getAppSettings();

  if( accountData.success === true && appSettingsData.success === true ){

    const plaidBalanceHistoryEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/accounts/balance/get';
  
    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      access_token: accountData.result.access_token,
      options: {
        account_ids : [ account_id ]
      }
    };
    response = plaidRequest(plaidBalanceHistoryEndpoint, payload);
  }
  return response;
}

function getPlaidInvestmentsData( account_id ){

  let response = { error: true, error_message: 'Account or settings lookup failed' };
  const accountData = getAppPlaidAccountById(account_id);
  const appSettingsData = getAppSettings();

  if( accountData.success === true && appSettingsData.success === true ){

    const plaidInvestmentHistoryEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/investments/holdings/get';
    
    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      access_token: accountData.result.access_token,
      options: {
        account_ids : [ account_id ]
      }
    };

    response = plaidRequest(plaidInvestmentHistoryEndpoint, payload);
  }
  return response;
}

function getPlaidItem(access_token){
  let result = [];
  const appSettingsData = getAppSettings();
  if( appSettingsData.success === true ){
    const plaidItemEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/item/get';
    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      access_token: access_token
    };
    let response = plaidRequest(plaidItemEndpoint, payload);
    return response.item;
  }
  return result;
}


function plaidRequest(url, payload) {
  var responseJson = { error: true, error_message: 'Unknown error' };
  var success = false;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) {
      Logger.log('Plaid API error: ' + JSON.stringify(json));
      handlePlaidError(json.error_type, json.error_code, json.error_message);
      responseJson = { error: true, error_type: json.error_type, error_code: json.error_code, error_message: json.error_message };
    } else {
      success = true;
      responseJson = json;
    }
  } catch (e) {
    Logger.log('plaidRequest exception: ' + e.toString());
    responseJson = { error: true, error_message: e.toString() };
  }
  logPlaidApiUsage_(url, getPlaidEndpointType_(url), success);
  return responseJson;
}

/**
 * Derives a readable endpoint_type from a Plaid URL path.
 * e.g. "https://sandbox.plaid.com/transactions/sync" → "transactions_sync"
 */
function getPlaidEndpointType_(url) {
  try {
    var path = url.replace(/^https?:\/\/[^\/]+/, ''); // strip scheme + host
    path = path.replace(/^\/+/, '').replace(/\/+$/, ''); // strip leading/trailing slashes
    return path.replace(/\//g, '_');
  } catch (e) {
    return 'unknown';
  }
}

/**
 * Fires a non-blocking usage log to the backend.
 * Failures are silently swallowed so they never interrupt the Plaid flow.
 */
function logPlaidApiUsage_(plaidUrl, billingType, status) {
  try {
    var email = '';
    try { email = UserEmail || ''; } catch(e) {}
    if (!email) {
      try { email = PropertiesService.getUserProperties().getProperty('USER_EMAIL') || ''; } catch(e) {}
    }

    var payload = {
      email: email,
      billing: billingType,
      endpoint: plaidUrl,
      status: status,
      timestamp: new Date().toISOString()
    };

    UrlFetchApp.fetch(API_ENDPOINT + 'api/plaid/usage/log', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: getAuthHeaders()
    });
    
  } catch (e) {
    Logger.log('[USAGE-LOG] Failed to log Plaid API usage: ' + e.toString());
  }
}

function handlePlaidError(error_type, error_code, error_message) {
  Logger.log('Plaid ' + (error_type || 'UNKNOWN') + ' [' + (error_code || '') + ']: ' + (error_message || ''));
  try {
    SpreadsheetApp.getUi().alert(error_message || 'A Plaid API error occurred.');
  } catch (e) {
    // UI not available in trigger context
  }
}
