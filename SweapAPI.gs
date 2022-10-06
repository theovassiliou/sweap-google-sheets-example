/*
MIT License

Copyright (c) 2022 Theo Vassiliou

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

/* -- Modify the client credentials --- */
const clientId = "MYCLIENTID"
const clientSecret = "MYCLIENTSECRET"


/* ----- No need to modify anything below this line ---- */
const tokenUrl = "https://auth.sweap.io/realms/users/protocol/openid-connect/token"
const apiUrl = "https://api.sweap.io/core/v1/"

/**
 * Retrieves a list of guests for the given eventName using the Sweap API.
 *
 * @param {"Demo Event"} eventName The name of the event as returned by ListEvents()
 * @return {Information} returns the LastName, FirstName, InvitationState, EntourageCount, AttendanceState for each guest.
 * @customfunction
 */
function ListGuests(eventName) {

  if ((eventName == undefined) || (eventName == "")) {
    return "No event name provided"
  }
    const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell();
    const guestResponseObject = JSON.parse(sweapGuests(eventName));

  if (guestResponseObject.length < 1) {
    return "No guests found!"
  }
    var rows = []

     for(var i = 0; i < guestResponseObject.length ; i++) {
      var cols = []
       cols.push(guestResponseObject[i].lastName);
       cols.push(guestResponseObject[i].firstName);
       cols.push(guestResponseObject[i].invitationState);
       cols.push(guestResponseObject[i].entourageCount);
       cols.push(guestResponseObject[i].attendanceState);
      rows.push(cols)
    }

    return rows;}

/**
 * Retrieves a list of events accesible to the user for the given eventName using the Sweap API. 
 * Returns all events if eventName is not given.
 *
 * @param {"Demo Event"} eventName Optional Filters the list of events for eventName. 
 * @return returns the Event Name, Start Date, End Date, State for each event
 * @customfunction
 */

function ListEvents(eventName) {
    // Check the current cell to detect recalculations due to reopening the sheet
    const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell();
    const responseObject = JSON.parse(sweapEvents(eventName));

    if (responseObject.length < 1) {
      return "Event \""+ eventName + "\" not found"
    }

    var rows = []

     for(var i = 0; i < responseObject.length ; i++) {
      var cols = []
       cols.push(responseObject[i].name);
       cols.push(new Date(responseObject[i].startDate));
       cols.push(new Date(responseObject[i].endDate));
        cols.push(responseObject[i].state);
      rows.push(cols)
    }

    
    return rows;
}
/* --- helper functions --- */
function sweapEvents(eventName) {
  const authUrlFetch = OAuth2.withClientCredentials(tokenUrl, clientId, clientSecret);
  if ((eventName != undefined) && (eventName != "")) {
    params = "?name="+eventName
  } else {
    params = ""
  }

  const response = authUrlFetch.fetch(apiUrl+"events"+params);
  checkResponse_(response);
  return response.getContentText();
}

function sweapGuestsByEventId(eventId) {
  const authUrlFetch = OAuth2.withClientCredentials(tokenUrl, clientId, clientSecret);
  const response = authUrlFetch.fetch(apiUrl+"guests?event_id="+eventId);
  checkResponse_(response);
  return response.getContentText();
}

function sweapGuests(eventName) {
  const responseObject = JSON.parse(sweapEvents(eventName));

  if (responseObject.length < 1) {
    return "Event not found for guests!"
  }

  theEventId = responseObject[0].id;
  return sweapGuestsByEventId(theEventId)

}


/**
 * Simple library for sending OAuth2 authenticated requests.
 * See: https://developers.google.com/google-ads/scripts/docs/features/third-party-apis#oauth_2
 * for full details.
 */

/**
 * Adds a OAuth object, for creating authenticated requests, to the global
 * object.
 */
(function(scope) {
  /**
   * Creates an object for making authenticated URL fetch requests with a
   * given stored access token.
   * @param {string} accessToken The access token to store and use.
   * @constructor
   */
  function OAuth2UrlFetchApp(accessToken) { this.accessToken_ = accessToken; }

  /**
   * Performs an HTTP request for the given URL.
   * @param {string} url The URL to fetch
   * @param {?Object=} options Options as per UrlFetchApp.fetch
   * @return {!HTTPResponse} The HTTP Response object.
   */
  OAuth2UrlFetchApp.prototype.fetch = function(url, opt_options) {
    const fetchOptions = opt_options || {};
    if (!fetchOptions.headers) {
      fetchOptions.headers = {};
    }
    fetchOptions.headers.Authorization = 'Bearer ' + this.accessToken_;
    return UrlFetchApp.fetch(url, fetchOptions);
  };

  /**
   * Performs the authentication step
   * @param {string} tokenUrl The endpoint for use in obtaining the token.
   * @param {!Object} payload The authentication payload, typically containing
   *     details of the grant type, credentials etc.
   * @param {string=} opt_authHeader Client credential grant also can make use
   *     of an Authorisation header, as specified here
   * @param {string=} opt_scope Optional string of spaced-delimited scopes.
   * @return {string} The access token
   */
  function authenticate_(tokenUrl, payload, opt_authHeader, opt_scope) {
    const options = {muteHttpExceptions: true, method: 'POST', payload: payload};
    if (opt_scope) {
      options.payload.scope = opt_scope;
    }
    if (opt_authHeader) {
      options.headers = {Authorization: opt_authHeader};
    }
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseData = JSON.parse(response.getContentText());
    if (responseData && responseData.access_token) {
      const accessToken = responseData.access_token;
      return accessToken;
    } else {
      throw Error('No access token received: ' + response.getContentText());
    }
    
  }

  /**
   * Creates a OAuth2UrlFetchApp object having authenticated with a refresh
   * token.
   * @param {string} tokenUrl The endpoint for use in obtaining the token.
   * @param {string} clientId The client ID representing the application.
   * @param {string} clientSecret The client secret.
   * @param {string} refreshToken The refresh token obtained through previous
   *     (possibly interactive) authentication.
   * @param {string=} opt_scope Space-delimited set of scopes.
   * @return {!OAuth2UrlFetchApp} The object for making authenticated requests.
   */
  function withRefreshToken(
      tokenUrl, clientId, clientSecret, refreshToken, opt_scope) {
    const payload = {
      grant_type: 'refresh_token',
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken
    };
    const accessToken = authenticate_(tokenUrl, payload, null, opt_scope);
    return new OAuth2UrlFetchApp(accessToken);
  }

  /**
   * Creates a OAuth2UrlFetchApp object having authenticated with client
   * credentials.
   * @param {string} tokenUrl The endpoint for use in obtaining the token.
   * @param {string} clientId The client ID representing the application.
   * @param {string} clientSecret The client secret.
   * @param {string=} opt_scope Space-delimited set of scopes.
   * @return {!OAuth2UrlFetchApp} The object for making authenticated requests.
   */
  function withClientCredentials(tokenUrl, clientId, clientSecret, opt_scope) {
    const authHeader =
        'Basic ' + Utilities.base64Encode([clientId, clientSecret].join(':'));
    const payload = {
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret
    };
    const accessToken = authenticate_(tokenUrl, payload, authHeader, opt_scope);
    return new OAuth2UrlFetchApp(accessToken);
  }

  /**
   * Creates a OAuth2UrlFetchApp object having authenticated with OAuth2 username
   * and password.
   * @param {string} tokenUrl The endpoint for use in obtaining the token.
   * @param {string} clientId The client ID representing the application.
   * @param {string} username OAuth2 Username
   * @param {string} password OAuth2 password
   * @param {string=} opt_scope Space-delimited set of scopes.
   * @return {!OAuth2UrlFetchApp} The object for making authenticated requests.
   */
  function withPassword(tokenUrl, clientId, username, password, opt_scope) {
    const payload = {
      grant_type: 'password',
      client_id: clientId,
      username: username,
      password: password
    };
    const accessToken = authenticate_(tokenUrl, payload, null, opt_scope);
    return new OAuth2UrlFetchApp(accessToken);
  }

  /**
   * Creates a OAuth2UrlFetchApp object having authenticated as a Service
   * Account.
   * Flow details taken from:
   *     https://developers.google.com/identity/protocols/OAuth2ServiceAccount
   * @param {string} tokenUrl The endpoint for use in obtaining the token.
   * @param {string} serviceAccount The email address of the Service Account.
   * @param {string} key The key taken from the downloaded JSON file.
   * @param {string} scope Space-delimited set of scopes.
   * @return {!OAuth2UrlFetchApp} The object for making authenticated requests.
   */
  function withServiceAccount(tokenUrl, serviceAccount, key, scope) {
    const assertionTime = new Date();
    const jwtHeader = '{"alg":"RS256","typ":"JWT"}';
    const jwtClaimSet = {
      iss: serviceAccount,
      scope: scope,
      aud: tokenUrl,
      exp: Math.round(assertionTime.getTime() / 1000 + 3600),
      iat: Math.round(assertionTime.getTime() / 1000)
    };
    const jwtAssertion = Utilities.base64EncodeWebSafe(jwtHeader) + '.' +
        Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
    const signature = Utilities.computeRsaSha256Signature(jwtAssertion, key);
    jwtAssertion += '.' + Utilities.base64Encode(signature);
    const payload = {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwtAssertion
    };
    const accessToken = authenticate_(tokenUrl, payload, null);
    return new OAuth2UrlFetchApp(accessToken);
  }

  scope.OAuth2 = {
    withRefreshToken: withRefreshToken,
    withClientCredentials: withClientCredentials,
    withServiceAccount: withServiceAccount,
    withPassword: withPassword
  };
})(this);

/**
 * Helper function to check response code is OK and if not, throw useful exceptions.
 */
function checkResponse_(response) {
    const responseCode = response.getResponseCode();
    if (200 <= responseCode && responseCode < 400) return;

    const content = response.getContentText();

    let message = '';
    try {
        const jsonObj = JSON.parse(content);
        if (jsonObj.message !== undefined) {
            message += `, message: ${jsonObj.message}`;
        }
        if (jsonObj.detail !== undefined) {
            message += `, detail: ${jsonObj.detail}`;
        }
    } catch (error) {
        // JSON parsing errors are ignored, and we fall back to the raw content
        message = ', ' + content;
    }

    switch (responseCode) {
        case 403:
            throw new Error(`Authorization failure, check authKey${message}`);
        case 456:
            throw new Error(`Quota for this billing period has been exceeded${message}`);
        case 400:
            throw new Error(`Bad request${message}`);
        case 429:
            throw new Error(
                `Too many requests, DeepL servers are currently experiencing high load${message}`,
            );
        default: {
            throw new Error(
                `Unexpected status code: ${responseCode} ${message}, content: ${content}`,
            );
        }
    }
}