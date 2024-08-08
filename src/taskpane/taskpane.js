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
async function initialize() {
  if (Office.context.mailbox && Office.context.mailbox.item) {
    const item = Office.context.mailbox.item;
    const subject = getLocationCode(item.subject);

    if (subject) {
      const locationCode = subject || 'NE1075';

      // Get user identity token
      const token = await getUserIdentityToken();

      // Send the token to ServiceNow to establish the session
      await establishServiceNowSession(token);

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
      previewElement.appendChild(iframe);
    }
  }
}

function getLocationCode(input) {
  const parts = input.split(' - ');
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

async function establishServiceNowSession(token) {
  try {


    const locationCode = 'NE1075';
    console.log('locationCode', locationCode)
    if (locationCode) {
      var response = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode)
      console.log('JRBP -> response:', response);
      var data = response.data?.result;
      console.log('>>>>> 1 ', data[0]);
      if (data && data[0]) {
        var sys_id = data[0].sys_id
        window.location.reload();
        var el = document.createElement("iframe");
        el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?search=' + sys_id;
        el.id = 'miIframe';
        el.referrerpolicy = "strict-origin-when-cross-origin";
        var a = document.getElementById("miIframe")?.remove();
        document.getElementById("preview").appendChild(el);
        const item = Office.context.mailbox.item;
      }
    }

    /*
        const response = await fetch('https://iadbdev.service-now.com/api/x_nuvo_eam_authentication/session', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`  // Send the token in the header
          },
          body: JSON.stringify({ token: token })
        });

        if (!response.ok) {
          throw new Error('Failed to establish session with ServiceNow');
        }

        const data = await response.json();
        console.log('Session established with ServiceNow:', data);*/
  } catch (error) {
    console.error('Error establishing session with ServiceNow:', error);
  }
}
