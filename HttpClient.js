function logStructuredEvent(eventName, details) {
  try {
    Logger.log(JSON.stringify({
      event: eventName,
      spreadsheetId: getUserSpreadsheetId(),
      timestamp: new Date().toISOString(),
      details: details || {},
    }));
  } catch (e) {
    Logger.log(eventName + ': ' + (e && e.toString ? e.toString() : e));
  }
}

const HTTP_CLIENT_CONFIG = {
  maxRetries: 3,
  backoffMs: 500,
};

function isRetryableStatusCode(statusCode) {
  return statusCode === 429 || (statusCode >= 500 && statusCode <= 599);
}

function sleepWithBackoff(attempt) {
  const waitMs = HTTP_CLIENT_CONFIG.backoffMs * Math.pow(2, attempt);
  Utilities.sleep(waitMs);
}

function requestJson(url, options) {
  const requestOptions = Object.assign({}, options || {}, {
    muteHttpExceptions: true,
  });

  const maxRetries = typeof requestOptions.maxRetries === 'number'
    ? requestOptions.maxRetries
    : HTTP_CLIENT_CONFIG.maxRetries;

  delete requestOptions.maxRetries;

  if (!requestOptions.method) {
    requestOptions.method = 'get';
  }

  let lastError = null;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, requestOptions);
      const statusCode = response.getResponseCode();
      const body = response.getContentText();

      let parsedBody = null;
      if (body) {
        try {
          parsedBody = JSON.parse(body);
        } catch (parseError) {
          parsedBody = { raw: body };
        }
      }

      if (isRetryableStatusCode(statusCode) && attempt < maxRetries) {
        logStructuredEvent('http.retry', { url: url, statusCode: statusCode, attempt: attempt + 1 });
        sleepWithBackoff(attempt);
        continue;
      }

      const ok = statusCode >= 200 && statusCode < 300;
      if (!ok) {
        logStructuredEvent('http.error', { url: url, statusCode: statusCode, responseBody: parsedBody });
      }

      return {
        success: ok,
        statusCode: statusCode,
        body: parsedBody,
        rawBody: body,
      };
    } catch (error) {
      lastError = error;
      if (attempt < maxRetries) {
        logStructuredEvent('http.exception.retry', { url: url, attempt: attempt + 1, error: error.toString() });
        sleepWithBackoff(attempt);
        continue;
      }
    }
  }

  logStructuredEvent('http.exception.final', { url: url, error: lastError ? lastError.toString() : 'Unknown request error' });

  return {
    success: false,
    statusCode: 0,
    body: null,
    rawBody: null,
    error: lastError ? lastError.toString() : 'Unknown request error',
  };
}
