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
  }
});
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
  try {
    var resp = await instance.post('/oauth_token.do?grant_type=password&client_id=f3600e11ee4bca94785814825f74d23a&client_secret=wUi%26mLGH0f&password=AutoCadIntegration67%3D&username=autocad_integration')
    console.log('JRBP -> resp:', resp);

  } catch (error) {
    console.log('JRBP -> error:', error);
  }
}