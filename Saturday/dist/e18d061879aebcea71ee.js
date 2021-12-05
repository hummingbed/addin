function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global console, localStorage, Office, require */
var Msal = require("msal");

Office.initialize = function () {
  if (Office.context.ui.messageParent) {
    userAgentApp.handleRedirectCallback(authCallback); // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
    // stored login data in localStorage. So a direct call of acquireTokenRedirect
    // causes the error "User login is required". Once the user is logged in successfully
    // the first time, msal data in localStorage will prevent this error from ever hap-
    // pening again; but the error must be blocked here, so that the user can login
    // successfully the first time. To do that, call loginRedirect first instead of
    // acquireTokenRedirect.

    if (localStorage.getItem("loggedIn") === "yes") {
      userAgentApp.acquireTokenRedirect(requestObj);
    } else {
      // This will login the user and then the (response.tokenType === "id_token")
      // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
      // and then the dialog is redirected back to this script, so the
      // acquireTokenRedirect above runs.
      userAgentApp.loginRedirect(requestObj);
    }
  }
};

var msalConfig = {
  auth: {
    clientId: "e3afd3e2-8b9f-42dc-a85f-aa193d499c07",
    //This is your client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000/fallbackauthdialog.html",
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage",
    // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true // Recommended to avoid certain IE/Edge issues.

  }
};
var requestObj = {
  scopes: ["https://graph.microsoft.com/User.Read"]
};
var userAgentApp = new Msal.UserAgentApplication(msalConfig);

function authCallback(error, response) {
  if (error) {
    console.log(error);
    Office.context.ui.messageParent(JSON.stringify({
      status: "failure",
      result: error
    }));
  } else {
    if (response.tokenType === "id_token") {
      console.log(response.idToken.rawIdToken);
      localStorage.setItem("loggedIn", "yes");
    } else {
      console.log("token type is:" + response.tokenType);
      Office.context.ui.messageParent(JSON.stringify({
        status: "success",
        result: response.accessToken
      }));
    }
  }
}

function onMessageComposeHandler(_x) {
  return _onMessageComposeHandler.apply(this, arguments);
}

function _onMessageComposeHandler() {
  _onMessageComposeHandler = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee(event) {
    return regeneratorRuntime.wrap(function _callee$(_context) {
      while (1) {
        switch (_context.prev = _context.next) {
          case 0:
            Office.context.mailbox.item.body.setAsync("temporary: setup for signature isn't available yet!!", {
              "asyncContext": event
            }, function (asyncResult) {
              // Handle success or error.
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
              } // Call event.completed() after all work is done.


              asyncResult.asyncContext.completed();
            });
            fetch("https://login.microsoftonline.com/d89e6156-6bbd-459b-9fb1-c866bb0a8c65/oauth2/v2.0/token", {
              method: "POST",
              mode: "no-cors",
              body: JSON.stringify({
                "client_id": "b1ccc2ff-45cf-437e-8ae2-5c44f0feea76",
                "scope": "https://graph.microsoft.com/.default offline_access",
                "client_secret": "~aJ7Q~FgT.ure3X~DXDm-xjC3R5RvFPfUBCp~",
                "grant_type": "client_credentials"
              })
            }).then(function (reposone) {
              console.log(reposone);
            });

          case 2:
          case "end":
            return _context.stop();
        }
      }
    }, _callee);
  }));
  return _onMessageComposeHandler.apply(this, arguments);
}

Office("onMessageComposeHandler", onMessageComposeHandler); // let tokenResponse = userAgentApp.acquireTokenSilent(requestObj)
// console.log('Login Response', tokenResponse);