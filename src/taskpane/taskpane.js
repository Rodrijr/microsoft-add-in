/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */
console.log('AAAAAAAAAAAAAFUERA 1');

const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com/api/',
  timeout: 1000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

console.log('AAAAAAAAAAAAAFUERA 2');

Office.onReady((info) => {
  console.log('info.host', info.host)
  console.log('Office.HostType.Outlook', Office.HostType.Outlook)
  if (info.host === Office.HostType.Outlook) {
  }
  console.log('Office.onReady')
  run();
});
function getLocationCode(input) {
  const parts = input.split(' - ');
  if (parts.length >= 2) {
    return parts[1];
  }
  return null;
}
async function run() {
  try {
    const item = Office.context.mailbox.item;
    const subject = item.subject;
    const locationCode = getLocationCode(subject);
    if (locationCode) {
      var { data } = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode)
      console.log('>>>>>', data[0]);
      if (data && data[0]) {
        var sys_id = data[0].sys_id
        var el = document.createElement("iframe");
        el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?search=' + sys_id;
        el.id = 'miIframe';
        el.referrerpolicy = "strict-origin-when-cross-origin";
        document.getElementById("miIframe")?.remove();
        document.getElementById("preview").appendChild(el);
        const item = Office.context.mailbox.item;
      }
    }
  } catch (error) {
    console.log('error >>>>>>>>>', error);
  }
}
run();
console.log('AAAAAAAAAAAAAFUERA');
