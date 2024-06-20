/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */
console.log('AAAAAAAAAAAAAFUERA 1')

const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com/api/',
  timeout: 1000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

console.log('AAAAAAAAAAAAAFUERA 2')

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    run();
  }
});

export async function run() {
  console.log('AAAAAAAAAAAAAAAAAQQQQQQQQQQQQQQQUIIIIIIIIIIIIIIII')
  instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=NE0C31')
    .then((response) => {
      console.log('>>>>>', response)
    }).catch(error => {
      console.log('error >>>>>>>>>', error)
    });

  var el = document.createElement("iframe");
  el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=60a9c6c31b4460504e9886e9cd4bcbe1&search=aca9c6c31b4460504e9886e9cd4bcbe0&view=default&label_size=8&qr=true';
  el.id = 'miIframe';
  el.referrerpolicy = "strict-origin-when-cross-origin";
  document.getElementById("miIframe")?.remove();
  document.getElementById("preview").appendChild(el);
  const item = Office.context.mailbox.item;
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}

console.log('AAAAAAAAAAAAAFUERA');
