/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();
  }
});
const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com/api/',
  timeout: 5000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

function subjectCB(result) {
  return result;
}
async function initialize() {

  if (Office.context.mailbox && Office.context.mailbox.item) {
    console.log('JRBP -> Office.context.mailbox:', Office.context.mailbox);
    console.log('JRBP ->  Office.context.mailbox.item:', Office.context.mailbox.item);
    let item = Office.context.mailbox.item;
    let sub = '';
    if (typeof item.subject == 'string') {
      sub = item.subject
    } else {
      sub = item.subject.getAsync()
      sub = sub.value;
    }

    console.log('JRBP -> sub:', sub);
    const subject = 'NE1075' || getLocationCode(sub);

    if (subject) {
      await establishServiceNowSession(subject);
    }
  }
}

function getLocationCode(input) {
  const parts = input?.split(' - ');
  return parts.length >= 2 ? parts[1] : null;
}

function getUserIdentityToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getUserIdentityTokenAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const token = result.value;
        console.log('User Identity Token:', token);
        Office.context.ui.messageParent(token);
        resolve(token);
      } else {
        console.error('Failed to get user identity token:', result.error);
        reject(result.error);
      }
    });
  });
}

async function establishServiceNowSession(locationCode) {
  if (locationCode) {
    try {
      getUserIdentityToken()

      var sys_id = await getLocationID(locationCode);
    } catch (error) {
      console.error('Error establishing session with ServiceNow:', error.response.status);
      console.error('Error establishing session with ServiceNow:', error.response.statusText);
      Office.context.ui.displayDialogAsync('https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id,
        { height: 45, width: 55 },
        function (asyncResult) {
          Office.context.ui.messageParent(token);
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            // Show an error message
            console.log('Failed to open dialog: ' + asyncResult.error.message);
          } else {
            getUserIdentityToken()
            var dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
              console.log('Message received from dialog: ' + args.message);
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
              console.log('Dialog closed: ' + args.error.message);
            });
          }
        });
      await getLocationID(locationCode);
    }
  }
}
async function getLocationID(locationCode) {
  var response = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode);
  console.log('JRBP -> response:', response);
  var data = response.data?.result;
  console.log('>>>>> 1 ', data[0]);
  if (data && data[0]) {

    var sys_id = data[0].sys_id;
    var el = document.createElement("iframe");
    el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id;
    console.log('JRBP -> el.src:', el.src);
    el.id = 'miIframe';
    el.sandbox = " allow-scripts allow-same-origin";
    el.referrerpolicy = "strict-origin-when-cross-origin";
    var a = document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);
    //window.location.replace('https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id);
    // document.location =
  }
  return sys_id;
}

Office.actions.associate("action", initialize);