/**
 * Retrieves user data and validates it against your external API.
 * Called from the sidebar when it loads.
 */
function validateUserSession() {

  try {
    const userEmail = getUserEmail();
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
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = requestJson(apiUrl, options);
    //Logger.log(result);
    return {
      success: response.success,
      result: response.result
    };

  } catch (e) {
    Logger.log( JSON.stringify(e, null, 2) );
    return { success: false, error: e.toString() };
  }
}

function getAppSettings() {
  try {
    const apiUrl = API_ENDPOINT + 'api/settings';

    const response = requestJson(apiUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
  
    return {
      success: response.success,
      result: response.result
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

function storePlaidAPIAccounts( data ){

  try {
    const userEmail = getUserEmail();
    
    // If email is empty, the user might need to re-authorize
    if (!userEmail) {
      throw new Error("Email access denied. Please re-authorize the add-on.");
    }
    // Your External API Endpoint
    const apiUrl = API_ENDPOINT + 'api/accounts/store-plaid';
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    const response = requestJson(apiUrl, options);
    //Logger.log(result);
    return {
      success: response.success,
      result: response.result
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAppPlaidConnectedAccounts() {

  try {

    const apiUrl = API_ENDPOINT + 'api/accounts/get-by-email/'+ getUserEmail();

    const options = {
      method: 'get',
      muteHttpExceptions: true
    };

    const response = requestJson(apiUrl, options);

    return {
      success: response.success,
      result: response.result
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
      muteHttpExceptions: true
    };

    const response = requestJson(apiUrl, options);

    return {
      success: response.success,
      result: response.result
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
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    const response = requestJson(apiUrl, options);

    return {
      success: response.success,
      result: response.result
    };

  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}

function requestJson(apiUrl, options) {
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseCode = response.getResponseCode();
  const contentText = response.getContentText();
  const success = responseCode >= 200 && responseCode < 300;

  if (!contentText) {
    return { success, result: null, responseCode };
  }

  try {
    const result = JSON.parse(contentText);
    const error = success
      ? null
      : result?.error || result?.message || `Request failed with status ${responseCode}.`;
    return { success, result, error, responseCode };
  } catch (error) {
    return {
      success: false,
      result: null,
      error: `Invalid JSON response (${responseCode}).`,
      responseCode
    };
  }
}
