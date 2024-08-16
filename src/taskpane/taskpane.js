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

// Obtener el token de autenticación básica
const authToken = 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=');

function subjectCB(result) {
  return result;
}

async function initialize() {
  if (Office.context.mailbox && Office.context.mailbox.item) {
    let item = Office.context.mailbox.item;
    let sub = '';

    if (typeof item.subject == 'string') {
      sub = item.subject;
    } else {
      sub = item.subject.getAsync().value;
    }

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
      await getUserIdentityToken();
      const sys_id = await getLocationID(locationCode);
    } catch (error) {
      console.error('Error establishing session with ServiceNow:', error.response.status);
      console.error('Error establishing session with ServiceNow:', error.response.statusText);

      Office.context.ui.displayDialogAsync('https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id,
        { height: 45, width: 55 },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log('Failed to open dialog: ' + asyncResult.error.message);
          } else {
            const dialog = asyncResult.value;
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
  const response = await instance.get(`now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=${locationCode}`);
  const data = response.data?.result;

  if (data && data[0]) {
    const sys_id = data[0].sys_id;
    const iframeUrl = `https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=${sys_id}`;

    // Aquí puedes abrir la página de ServiceNow en un iframe o en un diálogo modal
    openServiceNowPage(iframeUrl, authToken);

    return sys_id;
  }
}

function openServiceNowPage(url, authToken) {
  const el = document.createElement('iframe');
  el.src = url;
  el.id = 'miIframe';
  el.sandbox = 'allow-scripts allow-same-origin';
  el.referrerpolicy = 'strict-origin-when-cross-origin';
  el.setAttribute('Authorization', authToken);

  const preview = document.getElementById('preview');
  if (preview) {
    preview.innerHTML = '';
    preview.appendChild(el);
  } else {
    document.body.appendChild(el);
  }
}
