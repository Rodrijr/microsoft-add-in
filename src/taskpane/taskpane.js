/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();


  }
});



async function authenticateWithServiceNow() {
  const outlookAccessToken = getUserIdentityToken(); // Reemplaza con tu token actual

  try {
    const response = await fetch('https://iadbdev.service-now.com/api/now/table/sys_user', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${outlookAccessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        'user_name': 'autocad_integration',
        'password': 'AutoCadIntegration67=' // O algún otro mecanismo de autenticación
      })
    });

    const data = await response.json();
    console.log('JRBP -> data:', data);
    const serviceNowSessionToken = data.result.sys_id; // Suponiendo que el token está en el campo sys_id
    console.log('JRBP -> serviceNowSessionToken:', serviceNowSessionToken);

    // Cargar el iframe con el token de sesión
    const iframe = document.createElement('iframe');
    iframe.src = `https://iadbdev.service-now.com/nav_to.do?uri=x_nuvo_eam_microsoft_add_in.do&sysparm_session_id=${serviceNowSessionToken}`;
    var a = document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);
    // document.body.appendChild(iframe);
  } catch (error) {
    console.error('Error al autenticar con ServiceNow:', error);
  }
}



const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com/api/',
  timeout: 5000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

function subjectCB(result) {
  return result;
}

async function initialize() {
  // await getLocationID('NE1075');
  await authenticateWithServiceNow()
}

const authToken = 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=');

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
function sendToAddin() {
  // Suponiendo que el complemento está cargado en un iframe con el id 'myIframe'

 var accessToken = getUserIdentityToken();
  var iframe = document.getElementById('myIframe');
  iframe.contentWindow.postMessage({ type: 'accessToken', data: accessToken }, '*');
}
async function getLocationID(locationCode) {
  var response = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode);
  console.log('JRBP -> response:', response);
  var data = response.data?.result;
  console.log('>>>>> 1 ', data[0]);
  if (data && data[0]) {

    var sys_id = data[0].sys_id;
    var el = document.createElement("iframe");
    // https % 3A % 2F % 2Flo+
    // gin.microsoftonline.com % 2F9dfb1a05 - 5f1d - 449a - 8960 - 62abcb479e7d % 2Fsaml2 % 3FSAMLRequest % 3DlVLBjpswFPwV5DtgXEjACpHSRFUjbbdok % 252FbQm7EfWUtgUz % 252FDtn9f1km120O36tVv3sy8GW9QDD0b % 252BW7yj % 252BYBvk % 252BAPvox9Ab5dVKTyRluBWrkRgyA3Et % 252B2n264yyhfHTWW2l7Eu0QwXltzd4anAZwJ3CzlvDl4a4mj96PyNNUC9UqmBO8zmJjnxJph9SIeRQXSJQl0WFxoI14pnpZ7O1Fm2TQ0lm0nbem1wbCaqW6NhO0iIsuU3GeVyIuqxWNV0y0ss3XFaxVGk4h0QfrJIRLa9KJHoFEx0NNTvf7VVm1FDpWULZiqsoyKkrJaFuuSyhlni1AbASinuFlFXGCo0EvjK8JoyyPaRmz7JyVvKj4uzKhWfGNRM0to % 252FfaKG0ubwfaXkHIP57PTdx8Pp0DwawVuPsF % 252FX9ZfgWHIceFmmw3IQYefLvXJb9tSfxulmz % 252Fob1JXyvc9Eb % 252B7Pt4aGyv5c9o1 % 252Ff2ae9A % 252BOUW7yYItQzC % 252F91ElmThRau4C1A % 252BGRxB6k6DIun2JvvnJ97 % 252BAg % 253D % 253D % 26RelayState % 3Dhttps % 253A % 252F % 252Fiadbdev.service - now.com % 252Fsaml_redirector.do % 253Fsysparm_nostack % 253Dtrue % 2526sysparm_uri % 253D % 25252Fnav_to.do % 25253Furi % 25253D % 2525252Fx_nuvo_eam_fm_view_v2.do % 2525253Fapp % 2525253Duser
    el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id;
    console.log('JRBP -> el.src:', el.src);
    el.id = 'miIframe';
    el.sandbox = "allow-scripts allow-forms allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-modals allow-downloads allow-storage-access-by-user-activation";

    title = "Complemento de Office Locations finder"
    allow = ""
    el.name = "{&quot;baseFrameName&quot;:&quot;_xdm_5__8b7c90dc-80fd-0982-441d-9faa8998d12269854040_b7b7b150_1724265576652&quot;,&quot;hostInfo&quot;:&quot;Outlook|Web|16.01|es-ES|9c8bc367-191e-711c-b1d0-809b38add415|||16&quot;,&quot;xdmInfo&quot;:&quot;9edc182_813dbe25_1724265576652|8b7c90dc-80fd-0982-441d-9faa8998d122|https://outlook.office.com&quot;,&quot;flights&quot;:&quot;[\&quot;Microsoft.Office.SharedOnline.ProcessMultipleCommandsInDequeInvoker\&quot;]&quot;,&quot;disabledChangeGates&quot;:&quot;[]&quot;}"
    el.class = "AddinIframe"
    el.referrerpolicy = "strict-origin-when-cross-origin";
    //  el.setAttribute('Authorization', authToken);
    var a = document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);
    sendToAddin()
  } else {
    window.location.href = "https%3A%2F%2Flogin.microsoftonline.com%2F9dfb1a05-5f1d-449a-8960-62abcb479e7d%2Fsaml2%3FSAMLRequest%3DlVLBjpswFPwV5DtgXEjACpHSRFUjbbdok%252FbQm7EfWUtgUz%252FDtn9f1km120O36tVv3sy8GW9QDD0b%252BW7yj%252BYBvk%252BAPvox9Ab5dVKTyRluBWrkRgyA3Et%252B2n264yyhfHTWW2l7Eu0QwXltzd4anAZwJ3CzlvDl4a4mj96PyNNUC9UqmBO8zmJjnxJph9SIeRQXSJQl0WFxoI14pnpZ7O1Fm2TQ0lm0nbem1wbCaqW6NhO0iIsuU3GeVyIuqxWNV0y0ss3XFaxVGk4h0QfrJIRLa9KJHoFEx0NNTvf7VVm1FDpWULZiqsoyKkrJaFuuSyhlni1AbASinuFlFXGCo0EvjK8JoyyPaRmz7JyVvKj4uzKhWfGNRM0to%252FfaKG0ubwfaXkHIP57PTdx8Pp0DwawVuPsF%252FX9ZfgWHIceFmmw3IQYefLvXJb9tSfxulmz%252Fob1JXyvc9Eb%252B7Pt4aGyv5c9o1%252Ff2ae9A%252BOUW7yYItQzC%252F91ElmThRau4C1A%252BGRxB6k6DIun2JvvnJ97%252BAg%253D%253D%26RelayState%3Dhttps%253A%252F%252Fiadbdev.service-now.com%252Fsaml_redirector.do%253Fsysparm_nostack%253Dtrue%2526sysparm_uri%253D%25252Fnav_to.do%25253Furi%25253D%2525252Fx_nuvo_eam_fm_view_v2.do%2525253Fapp%2525253Duser"
  }
  return sys_id;
}
