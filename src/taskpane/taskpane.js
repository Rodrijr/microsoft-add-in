/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  var el = document.createElement("iframe");
  el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=60a9c6c31b4460504e9886e9cd4bcbe1&search=aca9c6c31b4460504e9886e9cd4bcbe0&view=default&label_size=8&qr=true';
  el.id = 'miIframe';
  document.getElementById("miIframe")?.remove();
  document.getElementById("preview").appendChild(el);



  const item = Office.context.mailbox.item;
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}
