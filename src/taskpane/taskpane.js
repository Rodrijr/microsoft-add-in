/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();
  }
});
//

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

      const locationCode = subject || 'NE1075';


      // Create an iframe and append it to the DOM
      iframe = document.createElement('iframe');
      iframe.src = "https://login.microsoftonline.com/9dfb1a05-5f1d-449a-8960-62abcb479e7d/saml2?SAMLRequest=lVLLbtswEPwVgXc9LVkSYRlwbRQ1kKZC7PbQ20pcOQQkUuVSSvv3VWgHSQ9N0St3dmZ2hhuCoU9Gvpvso3rAHxOS9X4OvSJ%2BnVRsMoprIElcwYDEbctPu893PAkiPhptdat75u2I0Fip1V4rmgY0JzSzbPHrw13FHq0diYehBNEInAO6znyln4JWD6GCeYQLBkIz77A4kAqeqV4Xe32RKhhkazTpzmrVS4VutRRdE0OU%2BVkXCz9NS%2FCLch356wSatknzEnMRulOY91GbFt2lFeugJ2Te8VCx0%2F2%2BgxLWkGVZUXbQJvlKrKIiyopCNDmkuFqAVAORnPF1lWjCoyILylYsiZLUjwo%2Fys9JwlcZj5MgLtLvzKtvGX2QSkh1eT%2FQ5goi%2Ful8rv36y%2BnsCGYp0Nwv6P%2FL8hsacjku1Gy7cTFw59u8Lfl9S%2FDSLNv%2BQ3sTvlW46Y382ffxUOtetr%2B8Xd%2Frp71BsMst1kzoahnA%2Ft1EHMTuRQq%2Fc1A%2BKRqxlZ1EwcLtTfbPT7z9DQ%3D%3D&RelayState=https%3A%2F%2Fiadbdev.service-now.com%2Fnavpage.do&sso_reload=true"
      iframe.id = 'miIframe';
      iframe.style.height = '100vh';
      iframe.style.width = '100vw';
      iframe.referrerpolicy = "strict-origin-when-cross-origin";

      // Append iframe to the preview element
      const previewElement = document.getElementById('preview');
      previewElement.innerHTML = ''; // Clear any existing content
      previewElement.appendChild(iframe);
      iframe = document.createElement('iframe');
      iframe.src = iframeUrl;
      iframe.id = 'miIframe';
      iframe.style.height = '100vh';
      iframe.style.width = '100vw';
      iframe.referrerpolicy = "strict-origin-when-cross-origin";

      // Append iframe to the preview element
      previewElement = document.getElementById('preview');
      previewElement.innerHTML = ''; // Clear any existing content
      previewElement.appendChild(iframe);
      // Not authenticated, prompt for login
    //  window.open('https://login.microsoftonline.com/', '_blank', 'width=500,height=600');
    }
  } catch (error) {
    console.error('Authentication check failed', error);
  }
}
