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
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    //Logger.log(result);
    return {
      success: response.getResponseCode() === 200,
      result
    };

  } catch (e) {
    Logger.log( JSON.stringify(e, null, 2) );
    return { success: false, error: e.toString() };
  }
}

function getAppSettings() {
  try {
    const apiUrl = API_ENDPOINT + 'api/settings';

    const options = {
      method: 'get',
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
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    //Logger.log(result);
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
