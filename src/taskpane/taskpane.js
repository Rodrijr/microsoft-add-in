/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();


  }
});

async function initialize() {
  console.log('>>>>>>>>>>>>>>>>>>>>>>>initialize')
  const el = document.createElement('iframe');
  el.src = "https://iadbdev.service-now.com/login.do";
  el.id = 'miIframe';
  el.sandbox = "allow-scripts allow-forms allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-modals allow-downloads allow-storage-access-by-user-activation";
  el.className = "AddinIframe";

  el.onload = function () {
    const iframeWindow = el.contentWindow;
    const iframeDocument = el.contentDocument || iframeWindow.document;
    console.log('>>>>>>>>>>>>>>>>>>> i frame on load')
    if (iframeWindow.location.href.includes("/login.do")) {
      const user = iframeDocument.getElementById("user_name");
      user.value = 'autocad_integration';

      const pass = iframeDocument.getElementById("user_password");
      pass.value = 'AutoCadIntegration67=';

      const button = iframeDocument.getElementById("sysverb_login");
      button.click();

      iframeWindow.location.href = "https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=";
    }
  };
  document.getElementById("miIframe")?.remove();
  document.getElementById("preview").appendChild(el);

// await authenticateWithServiceNow()
}