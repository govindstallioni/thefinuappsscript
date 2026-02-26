/**
 * Builds authenticated headers for backend API requests.
 * The backend should validate the OAuth token via Google's tokeninfo endpoint.
 */
function getAuthHeaders() {
  return {
    'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
    'X-User-Email': UserEmail || Session.getActiveUser().getEmail()
  };
}

/**
 * Retrieves user data and validates it against your external API.
 * Called from the sidebar when it loads.
 */
function validateUserSession() {

  try {
    const userEmail = UserEmail;
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    // If email is empty, the user might need to re-authorize
    if (!userEmail) {
      throw new Error("Email access denied. Please re-authorize the add-on.");
    }

    // Your External API Endpoint
    const apiUrl = API_ENDPOINT + 'api/users/validate-user';
    
    const payload = {
      email: userEmail,
      spreadsheetId: spreadsheetId,
      timestamp: new Date().toISOString()
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: getAuthHeaders(),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    // Normalize subscriptionId (accept common variants) so callers can reliably read it
    try{
      if(result && result.data){
        const d = result.data;
        const found = d.subscriptionId || d.subscription_id || (d.subscriptions && d.subscriptions.id) || (d.subscription && d.subscription.id) || d.subId || null;
        if(found){
          d.subscriptionId = found;
        }
      }
    }catch(e){
      // ignore normalization errors
    }
    return {
      success: response.getResponseCode() === 200,
      result
    };

  } catch (e) {
    Logger.log( JSON.stringify(e, null, 2) );
    return { success: false, error: e.toString() };
  }
}

var _cachedAppSettings = null;

function getAppSettings() {
  if (_cachedAppSettings) return _cachedAppSettings;
  try {
    const apiUrl = API_ENDPOINT + 'api/settings';

    const options = {
      method: 'get',
      headers: getAuthHeaders(),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());

    _cachedAppSettings = {
      success: response.getResponseCode() === 200,
      result
    };
    return _cachedAppSettings;

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

function clearAppSettingsCache() {
  _cachedAppSettings = null;
}

function storePlaidAPIAccounts( data ){

  try {
    const userEmail = UserEmail;
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    // If email is empty, the user might need to re-authorize
    if (!userEmail) {
      throw new Error("Email access denied. Please re-authorize the add-on.");
    }
    // Your External API Endpoint
    const apiUrl = API_ENDPOINT + 'api/accounts/store-plaid';
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: getAuthHeaders(),
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    return {
      success: response.getResponseCode() === 200,
      result
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAppPlaidConnectedAccounts() {

  try {

    const apiUrl = API_ENDPOINT + 'api/accounts/get-by-email/'+ UserEmail;

    const options = {
      method: 'get',
      headers: getAuthHeaders(),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());

    return {
      success: response.getResponseCode() === 200,
      result
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

function getAppPlaidAccountById( accountId ){
  try {

    const apiUrl = API_ENDPOINT + 'api/accounts/'+ accountId;

    const options = {
      method: 'get',
      headers: getAuthHeaders(),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());

    return {
      success: response.getResponseCode() === 200,
      result
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

function updateAppAccountDetailById( accountId, data ){

  try {

    const apiUrl = API_ENDPOINT + 'api/accounts/update-account/'+ accountId;

    const options = {
      method: 'patch',
      contentType: 'application/json',
      headers: getAuthHeaders(),
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());

    return {
      success: response.getResponseCode() === 200,
      result
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Cancels the user's subscription by notifying the external API.
 * The API will receive the user's email in the payload.
 */
function confirmCancelUserSubscription(){
  try{
    const apiUrl = API_ENDPOINT + 'api/payment/unsubscribe';
    const payload = {
      email: UserEmail,
      //spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
      //timestamp: new Date().toISOString()
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: getAuthHeaders(),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());

    // If API indicates success, clear local subscription progress
    if (response.getResponseCode() === 200) {
      try{ clearSubscriptionProgress(); }catch(e){}
    }

    return {
      success: response.getResponseCode() === 200,
      result
    };
  }catch(e){
    return { success: false, error: e.toString() };
  }
}
