async function generatePlaidTokenLink( access_token = null ) {
  
  let response = [];
  const appSettingsData = getAppSettings();

  if( appSettingsData.success === true ){

    const plaidEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/link/token/create';

    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      user: APP_USER_ID,
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
  //Logger.log( JSON.stringify(response, null, 2) );
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
    //Logger.log( JSON.stringify(dataToSend, null, 2) );
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

/**
 * Optional: Get accounts & institution info
 */
/*function getAccountsAndInstitution(access_token) {

  try{

    const appSettingsData = getAppSettings();

    if( appSettingsData.success === true ){

      const plaidAccountEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/accounts/get';
      const plaidInstitutionEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/institutions/get_by_id';

      // Get accounts
      const accRes = UrlFetchApp.fetch(plaidAccountEndpoint, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({
          client_id: appSettingsData.result.plaidClientKey,
          secret: appSettingsData.result.plaidSecretKey,
          access_token: access_token
        })
      });

      const accJson = JSON.parse(accRes.getContentText());

      // Get institution
      let instName = "Unknown";
      if (accJson.item?.institution_id) {
        const instRes = UrlFetchApp.fetch(plaidInstitutionEndpoint, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({
            client_id: appSettingsData.result.plaidClientKey,
            secret: appSettingsData.result.plaidSecretKey,
            institution_id: accJson.item.institution_id,
            //country_codes: PLAID_COUNTRY_CODES
          })
        });
        const instJson = JSON.parse(instRes.getContentText());
        instName = instJson.institution?.name || "Unknown";
      }

      return {
        accounts: accJson.accounts || [],
        institution: {
          id: accJson.item?.institution_id,
          name: instName
        }
      };
    }else{
      return {
        success: false,
        data: []
      };
    }
  }catch(error){
    return {
      success: false,
      error: error.toString()
    };
  }
}*/

function getPlaidTransactionSyncData( account_id, new_cursor = null ){

  let response = [];
  const accountData = getAppPlaidAccountById(account_id);
  const appSettingsData = getAppSettings();
  if( accountData.success === true && appSettingsData.success === true ){
    const plaidTransactionsEndpoint = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/transactions/sync';
    let payload = {
      client_id: appSettingsData.result.plaidClientKey,
      secret: appSettingsData.result.plaidSecretKey,
      access_token: accountData.result.access_token,
      cursor: new_cursor ? new_cursor : '',
      count: 500,
      options: {
        account_id : account_id
      }
    };
    response = plaidRequest(plaidTransactionsEndpoint, payload);
  }
  return response;
}

function getPlaidAccountBalance( account_id ){

  let response = [];
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

  let response = [];
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
  let response = [];
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
  return response;
}


function plaidRequest(url, payload) {
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText());
  } catch (e) {
    const body = JSON.parse(e.message.match(/\{.*\}/)?.[0] || '{}');
    const {error_type, error_code, error_message} = body;
    switch (error_type) {
      case 'ITEM_ERROR':
        handlePlaidItemError(error_code, error_message);
        break;
      case 'RATE_LIMIT_EXCEEDED':
        handleRateLimitError(error_code, error_message);
        break;
      case 'API_ERROR':
        handlePlaidAPIError(error_code, error_message);
        break;
      case 'INVALID_REQUEST':
      case 'INVALID_INPUT':
        handlePlaidInvalidError(error_code, error_message);
        break;
      case 'INSTITUTION_ERROR':
        handlePlaidInstitutionError(error_code, error_message);
        break;
      default:
        handlePlaidUnknownError(error_code, error_message);
        break;
    }
  }
}

function handlePlaidItemError( error_code, error_message ){
  SpreadsheetApp.getUi().alert(error_message);
}

function handleRateLimitError( error_code, error_message ){
  SpreadsheetApp.getUi().alert(error_message);

}

function handlePlaidAPIError( error_code, error_message ){
  SpreadsheetApp.getUi().alert(error_message);
}

function handlePlaidInvalidError( error_code, error_message ){
  SpreadsheetApp.getUi().alert(error_message);
}

function handlePlaidInstitutionError( error_code, error_message ){
  SpreadsheetApp.getUi().alert(error_message);
}

function handlePlaidUnknownError( error_code, error_message ){
  SpreadsheetApp.getUi().alert(error_message);
}
