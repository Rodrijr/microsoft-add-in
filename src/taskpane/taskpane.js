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
var subject;
Office.onReady((info) => {
  console.log('info.host', info.host)
  console.log('Office.HostType.Outlook', Office.HostType.Outlook)
  if (info.host === Office.HostType.Outlook) {

  }
  if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;
    console.log('item.subject' + item.subject)
    subject = getLocationCode(item.subject);

  }
  console.log('Office.onReady')
  run();
});
function getLocationCode(input) {
  console.log('>>' + input)
  const parts = input.split(' - ');
  if (parts.length >= 2) {
    console.log('>> parts:' + parts[1])
    return parts[1];
  }
  return null;
}
async function action() {
  try {
    const locationCode = subject ? subject : 'NE1075';
    console.log('locationCode', locationCode)
    if (locationCode) {
      var response= await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode)
      console.log('JRBP -> response:', response);
      var data = response.data?.result;
      console.log('>>>>> 1 ', data[0]);
      if (data && data[0]) {
        var sys_id = data[0].sys_id
        var el = document.createElement("iframe");
        el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?search=' + sys_id;
       // Office.context.ui.displayDialogAsync(el.src, { height: 70, width: 80 });
        el.id = 'miIframe';
        el.referrerpolicy = "strict-origin-when-cross-origin";
        var a = document.getElementById("miIframe")?.remove();
        document.getElementById("preview").appendChild(el);
       // const item = Office.context.mailbox.item;
      }
    }
  } catch (error) {
    console.log('error >>>>>>>>>', error);
  }
}
action();
console.log('AAAAAAAAAAAAAFUERA');
