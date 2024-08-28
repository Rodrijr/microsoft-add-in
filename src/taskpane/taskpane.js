/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();


  }
});

async function initialize() {
  const serviceNowTokenEndpoint = 'https://iadbdev.service-now.com/oauth_token.do';
  const clientId = process.env.TESTID;
  console.log('JRBP -> clientId:', clientId);
  const clientSecret = process.env.TEST;
  console.log('JRBP -> clientSecret:', clientSecret);
  const redirectUri = 'https://rodrijr.github.io/microsoft-add-in/src/taskpane/taskpane.html';

  fetch(serviceNowTokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body: `grant_type=authorization_code&client_id=${clientId}&client_secret=${clientSecret}&code=${accessToken}&redirect_uri=${redirectUri}`

  })
    .then(response => response.json())
    .then(data => {
      const serviceNowAccessToken = data.access_token;
      const iframeUrl = `https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=&sysparm_access_token=${serviceNowAccessToken}`;
      var el = document.createElement("iframe");
      el.src = iframeUrl;
      el.id = 'miIframe';
      el.sandbox = "allow-scripts allow-same-origin allow-top-navigation-by-user-activation";
      el.referrerpolicy = "strict-origin-when-cross-origin";
      el.setAttribute('Authorization', authToken);
      var a = document.getElementById("miIframe")?.remove();
      document.getElementById("preview").appendChild(el);
    })
    .catch(error => {
      console.error('Error al obtener el token de ServiceNow:', error);
    });
}