/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, require */
var ssoAuthHelper = require("./../helpers/ssoauthhelper");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;
  }
});