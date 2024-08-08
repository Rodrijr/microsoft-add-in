/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();
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

      // Authenticate the user in ServiceNow
      const authenticated = await authenticateUserInServiceNow(token);

      if (authenticated) {
        // Once authenticated, load the iframe with the restricted page
        loadIframe(locationCode);
      } else {
        console.error('User authentication failed.');
      }
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

async function authenticateUserInServiceNow(token) {
  try {
    const response = await axios.post('https://iadbdev.service-now.com/api/now/v1/session', { token });

    if (response.status === 200) {
      console.log('User authenticated in ServiceNow.');
      return true;
    } else {
      console.error('Authentication in ServiceNow failed:', response.statusText);
      return false;
    }
  } catch (error) {
    console.error('Error during authentication in ServiceNow:', error);
    return false;
  }
}

function loadIframe(locationCode) {
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
