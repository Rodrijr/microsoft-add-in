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
  location = 'https://iadbdev.service-now.com/login.do'

  if (info.host === Office.HostType.Outlook) {

    loginOAUTH();

  }
});

function redirectToPage() {

  location = 'https://iadbdev.service-now.com/x_nuvo_eam_fm_view_v2.do'
  console.log('JRBP -> location:', location);
}


function onloadHandler() {

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
  setTimeout(redirectToPage, 1000)
}
async function loginOAUTH() {
  try {

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
    console.log('JRBP -> loginOAUTH:', loginOAUTH);
  } catch (error) {
    console.log('JRBP -> error:', error);
  }
}