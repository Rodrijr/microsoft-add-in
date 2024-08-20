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
    'Authorization': 'Basic ' + btoa('user' + ':' + 'pass=')
  }
});

function subjectCB(result) {
  return result;
}

async function initialize() {
  await getLocationID(locationCode);

}

const authToken = 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=');

async function getLocationID(locationCode) {
  var response = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1&location_code=' + locationCode);
  console.log('JRBP -> response:', response);
  var data = response.data?.result;
  console.log('>>>>> 1 ', data[0]);
  if (data && data[0]) {

    var sys_id = data[0].sys_id;
    var el = document.createElement("iframe");
    el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do?app=user#?s=e2a369cd47dee5d08aba7f67536d4387&view=default&search=' + sys_id;
    console.log('JRBP -> el.src:', el.src);
    el.id = 'miIframe';
    el.sandbox = "allow-scripts allow-same-origin allow-top-navigation-by-user-activation";
    el.referrerpolicy = "strict-origin-when-cross-origin";
    el.setAttribute('Authorization', authToken);
    var a = document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);
  }
  return sys_id;
}
