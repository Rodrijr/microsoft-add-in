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
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

function subjectCB(result) {
  return result;
}

async function initialize() {
  await getLocationID('NE1075');

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
    el.sandbox = "allow-scripts allow-forms allow-same-origin allow-popups allow-popups-to-escape-sandbox allow-modals allow-downloads allow-storage-access-by-user-activation";

    title = "Complemento de Office Locations finder"
    allow = ""
    el.name = "{&quot;baseFrameName&quot;:&quot;_xdm_5__8b7c90dc-80fd-0982-441d-9faa8998d12269854040_b7b7b150_1724265576652&quot;,&quot;hostInfo&quot;:&quot;Outlook|Web|16.01|es-ES|9c8bc367-191e-711c-b1d0-809b38add415|||16&quot;,&quot;xdmInfo&quot;:&quot;9edc182_813dbe25_1724265576652|8b7c90dc-80fd-0982-441d-9faa8998d122|https://outlook.office.com&quot;,&quot;flights&quot;:&quot;[\&quot;Microsoft.Office.SharedOnline.ProcessMultipleCommandsInDequeInvoker\&quot;]&quot;,&quot;disabledChangeGates&quot;:&quot;[]&quot;}"
    el.class="AddinIframe"
    el.referrerpolicy = "strict-origin-when-cross-origin";
  //  el.setAttribute('Authorization', authToken);
    var a = document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);
  }
  return sys_id;
}
