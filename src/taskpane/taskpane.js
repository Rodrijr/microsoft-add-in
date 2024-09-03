/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';
const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com',
  timeout: 5000,
    headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + btoa('autocad_integration' + ':' + 'AutoCadIntegration67=')
  }
});
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loginOAUTH();
    redirectToPage();

  }
});

function redirectToPage() {
  location = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do'
}
function onloadHandler() {

  console.log("ESTOY EN EL onloadHandler?????????????????????????????????????")
  var b= this.getElementById('user_name')
  console.log(b)
  this.addEventListener("DOMContentLoaded", function (event) {
    console.log("cargo la pagina 2222222222222222222")
    var a = this.getElementById('user_name')
    console.log(a)
  })
}
document.addEventListener("DOMContentLoaded", function (event) {

  console.log("cargo la pagina 12111111111111111111")
})
async function loginOAUTH() {
  fetch("https://iadbdev.service-now.com/login.do", {
    "headers": {
      "content-type": "application/x-www-form-urlencoded",
      "sec-ch-ua": "\"Chromium\";v=\"128\", \"Not;A=Brand\";v=\"24\", \"Google Chrome\";v=\"128\"",
      "sec-ch-ua-mobile": "?0",
      "sec-ch-ua-platform": "\"Windows\"",
      "upgrade-insecure-requests": "1",
      "Referer": "https://iadbdev.service-now.com/login.do",
      "Referrer-Policy": "same-origin"
    },
    "body": "sysparm_ck=59d51e2f479452d46f0ee52f016d43e6853443e8b933c9c89a15a2e1084eba8bbf2668c7&user_name=autocad_integration&user_password=AutoCadIntegration67%3D&ni.nolog.user_password=true&ni.noecho.user_name=true&ni.noecho.user_password=true&language_select=en&screensize=1920x1080&sys_action=sysverb_login&not_important=",
    "method": "POST"
  });
}