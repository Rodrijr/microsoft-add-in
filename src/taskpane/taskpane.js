Office.onReady((info) => {
  getUserData()
  if (info.host === Office.HostType.Outlook) {
    if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      const item = Office.context.mailbox.item;
      subject = getLocationCode(item.subject);
    }
  }
  checkServiceNowSession();
});


async function getUserData() {
  try {
    let userTokenEncoded = await OfficeRuntime.auth.getAccessToken();
    let userToken = jwt_decode(userTokenEncoded); // Using the https://www.npmjs.com/package/jwt-decode library.
    console.log(">>>>>>>>>>>>>>>>>>> ", userToken.name); // user name
    console.log(">>>>>>>>>>>>>>>>>>> ", userToken.preferred_username); // email
    console.log(">>>>>>>>>>>>>>>>>>> ", userToken.oid); // user id
  }
  catch (exception) {
    if (exception.code === 13003) {
      // SSO is not supported for domain user accounts, only
      // Microsoft 365 Education or work account, or a Microsoft account.
    } else {
      // Handle error
    }
  }
}
function getLocationCode(input) {
  const parts = input.split(' - ');
  if (parts.length >= 2) {
    return parts[1];
  }
  return null;
}

async function checkServiceNowSession() {
  try {
    // Verificamos si la sesión de ServiceNow ya está abierta
    const response = await axios.get('https://iadbdev.service-now.com/api/now/v2/table/sys_user?sysparm_limit=1');
    if (response.status === 200) {
      // Si la sesión está activa, procedemos con la acción
      action();
    } else {
      // Si no, cargamos la página de autenticación dentro del iframe
      loadAuthPage();
    }
  } catch (error) {
    console.log('No active ServiceNow session found, loading auth page.');
    loadAuthPage();
  }
}

function loadAuthPage() {
  const iframe = document.getElementById("miIframe");
  iframe.src = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in_auth.do';
  iframe.id = 'miIframe';
  iframe.referrerpolicy = "strict-origin-when-cross-origin";
  document.getElementById("miIframe")?.remove();
  document.getElementById("preview").appendChild(iframe);
}

async function action() {
  try {
    const locationCode = subject ? subject : 'NE1075';
    if (locationCode) {
      const iframe = document.createElement("iframe");
      iframe.src = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=' + locationCode;
      iframe.id = 'miIframe';
      iframe.referrerpolicy = "strict-origin-when-cross-origin";
      document.getElementById("miIframe")?.remove();
      document.getElementById("preview").appendChild(iframe);
    }
  } catch (error) {
    console.log('Error in action:', error);
  }
}
