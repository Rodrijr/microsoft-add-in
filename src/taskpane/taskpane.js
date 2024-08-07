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
      const iframeUrl = `${locationEndpoint}${locationCode}`;

      // Create an iframe and append it to the DOM
      const iframe = document.createElement('iframe');
      iframe.src = iframeUrl;
      iframe.id = 'miIframe';
      iframe.style.height = '100vh';
      iframe.style.width = '100vw';
      iframe.referrerpolicy = "strict-origin-when-cross-origin";

      // Append iframe to the preview element
      const previewElement = document.getElementById('preview');
      previewElement.innerHTML = ''; // Clear any existing content
      previewElement.appendChild(iframe);

      // Check authentication without changing the session
      await checkAuthentication();
    }
  }
}

function getLocationCode(input) {
  const parts = input.split(' - ');
  return parts.length >= 2 ? parts[1] : null;
}

async function checkAuthentication() {
  try {
    // Attempt to load an authenticated page from ServiceNow
    const response = await fetch('https://iadbdev.service-now.com/api/now/table/sys_user', {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      credentials: 'include'
    });

    if (response.status === 401) {
      // Not authenticated, prompt for login
      window.open('https://login.microsoftonline.com/', '_blank', 'width=500,height=600');
    }
  } catch (error) {
    console.error('Authentication check failed', error);
  }
}
