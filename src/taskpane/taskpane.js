Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    if (Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      const item = Office.context.mailbox.item;
      subject = getLocationCode(item.subject);
    }
  }
  checkServiceNowSession();
});

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
      // Si no, redirigimos al usuario para autenticarse
      window.location.href = 'https://iadbdev.service-now.com';
    }
  } catch (error) {
    console.log('No active ServiceNow session found, redirecting to SSO login.');
    window.location.href = 'https://iadbdev.service-now.com';
  }
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
