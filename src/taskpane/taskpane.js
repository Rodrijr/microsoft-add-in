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
    //'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
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
      // const locationCode = subject || 'NE1075';
      // const locationCode = 'NE1075';

      // Send the token to ServiceNow to establish the session
      await establishServiceNowSession(subject);
      /*
            const iframeUrl = `${locationEndpoint}${locationCode}`;

            // Create an iframe and append it to the DOM
            const iframe = document.createElement('iframe');
            iframe.src = iframeUrl;
            iframe.id = 'miIframe';
            iframe.style.height = '100vh';
            iframe.style.width = '100vw';
            iframe.referrerpolicy = "strict-origin-when-cross-origin";

            const previewElement = document.getElementById('preview');
            previewElement.innerHTML = '';
            previewElement.appendChild(iframe);*/
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
        resolve(token);
      } else {
        console.error('Failed to get user identity token:', result.error);
        reject(result.error);
      }
    });
  });
}

async function establishServiceNowSession(locationCode) {
  console.log('JRBP ->=-a=dscasdcadscadsc asdc asdc locationCode:', locationCode);
  if (locationCode) {
    try {
      console.log('locationCode', locationCode)
      var response = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode)
      console.log('JRBP -> response:', response);
      var data = response.data?.result;
      console.log('>>>>> 1 ', data[0]);
      if (data && data[0]) {
        var sys_id = data[0].sys_id
        var el = document.createElement("iframe");
        el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id;
        console.log('JRBP -> el.src:', el.src);
        el.id = 'miIframe';
        el.sandbox = " allow-scripts allow-same-origin";
        el.referrerpolicy = "strict-origin-when-cross-origin";
        //var a = document.getElementById("miIframe")?.remove();
        //document.getElementById("preview").appendChild(el);
        window.location.replace('https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id);
      // document.location =
      }
    } catch (error) {
      console.error('Error establishing session with ServiceNow:', error);
    }
  }
}
