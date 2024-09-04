

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {

    var count = 0;
    var fetchFunction = async function () {
      console.log('HACER FETCH')

      if (location == 'https://iadbdev.service-now.com/login.do' && count < 5) {
        console.log('HACER FETCH 1')
        count++;
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
        location = 'https://iadbdev.service-now.com/login.do'

      }
    }
    setInterval(fetchFunction, 1000)
    setTimeout(function () {
      location ='https://iadbdev.service-now.com/login.do'
    }, 1500)
  }
});
