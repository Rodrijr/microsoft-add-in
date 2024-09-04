/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {

  }
});
async function onloadHandler() {
  try {
    console.log('>>>>>>>>>>>>>>>>>>>>>>>>>>> 1 ')
    console.log('LOADED');
    if (location == 'https://iadbdev.service-now.com/login.do') {

      var resp = await fetch("https://iadbdev.service-now.com/login.do", {
        "headers": {
          "content-type": "application/x-www-form-urlencoded",
          "sec-ch-ua": "\"Chromium\";v=\"128\", \"Not;A=Brand\";v=\"24\", \"Google Chrome\";v=\"128\"",
          "sec-ch-ua-mobile": "?0",
          "sec-ch-ua-platform": "\"Windows\"",
          "upgrade-insecure-requests": "1",
          "Referer": "https://rodrijr.github.io",
          "Referrer-Policy": "same-origin"
        },
        "body": "sysparm_ck=59d51e2f479452d46f0ee52f016d43e6853443e8b933c9c89a15a2e1084eba8bbf2668c7&user_name=autocad_integration&user_password=AutoCadIntegration67%3D&ni.nolog.user_password=true&ni.noecho.user_name=true&ni.noecho.user_password=true&language_select=en&screensize=1920x1080&sys_action=sysverb_login&not_important=",
        "method": "POST"
      });
      console.log('>>>>>>>>>>>>>>>>>>>>>>>>>>> 1 ', resp, location)
      console.log('>>>>>>>>>>>>>>>>>>>>>>>>>>> 1 ', document.location)
    } else {
      location = 'https://iadbdev.service-now.com/login.do';
      var resp = await fetch("https://iadbdev.service-now.com/login.do", {
        "headers": {
          "content-type": "application/x-www-form-urlencoded",
          "sec-ch-ua": "\"Chromium\";v=\"128\", \"Not;A=Brand\";v=\"24\", \"Google Chrome\";v=\"128\"",
          "sec-ch-ua-mobile": "?0",
          "sec-ch-ua-platform": "\"Windows\"",
          "upgrade-insecure-requests": "1",
          "Referer": "https://rodrijr.github.io",
          "Referrer-Policy": "same-origin"
        },
        "body": "sysparm_ck=59d51e2f479452d46f0ee52f016d43e6853443e8b933c9c89a15a2e1084eba8bbf2668c7&user_name=autocad_integration&user_password=AutoCadIntegration67%3D&ni.nolog.user_password=true&ni.noecho.user_name=true&ni.noecho.user_password=true&language_select=en&screensize=1920x1080&sys_action=sysverb_login&not_important=",
        "method": "POST"
      });
    }

    var iframe = document.getElementById('miIframe');
    //iframe.src = 'https://iadbdev.service-now.com/'
    window.frames["miIframe"].location = 'https://iadbdev.service-now.com/'
  } catch (error) {
    console.log('JRBP -> error:', error);
  }
}
/*
  function newFunction() {
    console.log('>>>>>>>>>>>>>>>>>>>>>>>>>>> REDIRECT TO: ' + 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=NE1081');
    window.location = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=NE1081';
  }
}
*/