/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {

    window.location = 'https://iadbdev.service-now.com/login.do';
   // setTimeout(loginOAUTH, 2000)
  }
});
/*async function loginOAUTH() {
  try {
    console.log('>>>>>>>>>>>>>>>>>>>>>>>>>>> 1 ')

    document.addEventListener("DOMContentLoaded", function (event) {
      console.log('LOADED');
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
      setTimeout(newFunction, 2000)

    })

  } catch (error) {
    console.log('JRBP -> error:', error);
  }

  function newFunction() {
    console.log('>>>>>>>>>>>>>>>>>>>>>>>>>>> REDIRECT TO: ' + 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=NE1081');
    window.location = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=NE1081';
  }
}
*/