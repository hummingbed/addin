function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global OfficeRuntime, require */
var documentHelper = require("./documentHelper");

var fallbackAuthHelper = require("./fallbackAuthHelper");

var sso = require("office-addin-sso");

var retryGetAccessToken = 0;
export function getGraphData() {
  return _getGraphData.apply(this, arguments);
}

function _getGraphData() {
  _getGraphData = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee() {
    var bootstrapToken, exchangeResponse, response;
    return regeneratorRuntime.wrap(function _callee$(_context) {
      while (1) {
        switch (_context.prev = _context.next) {
          case 0:
            _context.prev = 0;
            _context.next = 3;
            return OfficeRuntime.auth.getAccessToken({
              allowSignInPrompt: true
            });

          case 3:
            bootstrapToken = _context.sent;
            _context.next = 6;
            return sso.getGraphToken(bootstrapToken);

          case 6:
            exchangeResponse = _context.sent;

            if (exchangeResponse.claims) {// Microsoft Graph requires an additional form of authentication. Have the Office host
              // get a new token using the Claims string, which tells AAD to prompt the user for all
              // required forms of authentication.
              // let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
              // exchangeResponse = sso.getGraphToken(mfaBootstrapToken);
            }

            if (!exchangeResponse.error) {
              _context.next = 12;
              break;
            }

            // AAD errors are returned to the client with HTTP code 200, so they do not trigger
            // the catch block below.
            handleAADErrors(exchangeResponse);
            _context.next = 18;
            break;

          case 12:
            _context.next = 14;
            return sso.makeGraphApiCall(exchangeResponse.access_token);

          case 14:
            response = _context.sent;
            console.log('mwnwb', response); // console.log('old token', exchangeResponse.access_token)

            documentHelper.writeDataToOfficeDocument(response);
            sso.showMessage("Your data has been added to the document.");

          case 18:
            _context.next = 23;
            break;

          case 20:
            _context.prev = 20;
            _context.t0 = _context["catch"](0);

            if (_context.t0.code) {
              if (sso.handleClientSideErrors(_context.t0)) {
                fallbackAuthHelper.dialogFallback();
              }
            } else {
              sso.showMessage("EXCEPTION: " + JSON.stringify(_context.t0));
            }

          case 23:
          case "end":
            return _context.stop();
        }
      }
    }, _callee, null, [[0, 20]]);
  }));
  return _getGraphData.apply(this, arguments);
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    getGraphData();
  }
});

function handleAADErrors(exchangeResponse) {
  // On rare occasions the bootstrap token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired bootstrap token.
  if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1 && retryGetAccessToken <= 0) {
    retryGetAccessToken++;
    getGraphData();
  } else {
    fallbackAuthHelper.dialogFallback();
  }
}