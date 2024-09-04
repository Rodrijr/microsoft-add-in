// taskpane.js

window.onload = function () {
  // Notifica al taskpane que el iframe se ha cargado
  console.log('>>>>>>>>>>>>>>>>>>>>>>>> LOADED')
  window.parent.postMessage('iframe_loaded', '*');
};
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('>>>>>>>>>>>>>>>>>>>>>>>> LOADED 2222')

    window.parent.postMessage('iframe_loaded', '*');
  }
});

window.addEventListener('start_fetch', async function (event) {
    try {
      console.log('Iniciando fetch dentro del iframe...');
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

      console.log('>>>>>>>>>>>>>>>>>>>>>>>> Fetch completado:', resp);

      // Redirigir a otra página dentro del iframe
      window.location.href = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=NE1081';

    } catch (error) {
      console.log(' >>>>>>>>>>>>>>>>>>>>>>>> Error en el fetch:', error);
    }
});

function onloadIframe() {
  window.addEventListener('start_fetch', async function (event) {
    try {
      console.log('onloadIframe Iniciando fetch dentro del iframe...');
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

      console.log(' onloadIframe>>>>>>>>>>>>>>>>>>>>>>>> Fetch completado:', resp);

      // Redirigir a otra página dentro del iframe
      //window.location.href = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=NE1081';

    } catch (error) {
      console.log(' onloadIframe >>>>>>>>>>>>>>>>>>>>>>>> Error en el fetch:', error);
    }
  });

}