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
    // Assuming there's an API endpoint in ServiceNow to handle user authentication via token
    const response = await instance.post('now/v1/session', { token });

    if (response.status === 200) {
      console.log('User authenticated in ServiceNow.');
    } else {
      console.error('Failed to authenticate in ServiceNow:', response.statusText);
    }
  } catch (error) {
    console.error('Error establishing session with ServiceNow:', error);
  }
}
