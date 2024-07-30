/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */

const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com/api/',
  timeout: 1000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

var subject;
Office.onReady((info) => {
  console.log('info.host', info.host)
  console.log('Office.HostType.Outlook', Office.HostType.Outlook)
  if (info.host === Office.HostType.Outlook) {

  }
  if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;
    console.log('item.subject: ' + JSON.stringify(item))
    console.log('item.subject: ' ,(item))
    subject = getLocationCode(item.subject);

  }
  action();
});
function getLocationCode(input) {
  const parts = input.split(' - ');
  if (parts.length >= 2) {
    return parts[1];
  }
  return null;
}
async function action() {
  try {
    const locationCode = subject ? subject : 'NE1075';
    console.log('locationCode', locationCode)
    if (locationCode) {
      var el = document.createElement("iframe");
      el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=' + locationCode
      el.id = 'miIframe';
      el.referrerpolicy = "strict-origin-when-cross-origin";
      document.getElementById("miIframe")?.remove();
      document.getElementById("preview").appendChild(el);

    }
  } catch (error) {
    console.log('error', error);
  }
}
action();
