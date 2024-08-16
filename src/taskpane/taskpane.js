/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();
  }
});

async function initialize() {
  if (Office.context.mailbox && Office.context.mailbox.item) {
    let item = Office.context.mailbox.item;
    let sub = '';

    if (typeof item.subject === 'string') {
      sub = item.subject;
    } else {
      const asyncResult = await getSubjectAsync(item.subject);
      sub = asyncResult.value;
    }

    const subject = 'NE1075' || getLocationCode(sub);

    if (subject) {
      await establishServiceNowSession(subject);
    }
  }
}

function getSubjectAsync(subject) {
  return new Promise((resolve, reject) => {
    subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult);
      } else {
        reject(asyncResult.error);
      }
    });
  });
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
      // Obtener el token de identidad del usuario
      const userIdentityToken = await getUserIdentityToken();

      // Enviar el token a ServiceNow para autenticar la sesión
      await authenticateInServiceNow(userIdentityToken);

      // Hacer la llamada a la API de ServiceNow
      const response = await axios.get(`https://iadbdev.service-now.com/api/now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=${locationCode}`);
      const data = response.data?.result;

      if (data && data[0]) {
        const sys_id = data[0].sys_id;
        const iframeUrl = `https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=${sys_id}&view=default&search=${sys_id}`;

        window.location.replace(iframeUrl);
      }
    } catch (error) {
      console.error('Error establishing session with ServiceNow:', error);
    }
  }
}

async function authenticateInServiceNow(userIdentityToken) {
  try {
    const response = await axios.post('https://iadbdev.service-now.com/api/now/v1/sso_login', {
      token: userIdentityToken
    });

    if (response.status === 200) {
      console.log('User authenticated in ServiceNow');
    } else {
      console.error('Failed to authenticate in ServiceNow:', response.status, response.statusText);
    }
  } catch (error) {
    console.error('Error authenticating in ServiceNow:', error);
  }
}
