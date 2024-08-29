/* global Office */
const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com/api/',
  timeout: 1000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});

async function run() {
  try {
    var data = await instance.get('now/table/x_nuvo_eam_elocation?sysparm_fields=sys_id&sysparm_limit=1')
    console.log('>>>>>', data);

    var el = document.createElement("iframe");
    el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do';
    el.id = 'miIframe';
    el.referrerpolicy = "strict-origin-when-cross-origin";
    document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);

  } catch (error) {
    console.log('error >>>>>>>>>', error);
  }
}

run();