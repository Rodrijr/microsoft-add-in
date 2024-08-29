/* global Office */
const locationEndpoint = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=';
const instance = axios.create({
  baseURL: 'https://iadbdev.service-now.com',
  timeout: 5000,
  headers: {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*'
  }
});
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loginOAUTH();
  }
});

async function loginOAUTH() {
  try {
    var resp = await instance.post('/oauth_token.do?grant_type=password&client_id=f3600e11ee4bca94785814825f74d23a&client_secret=wUi%26mLGH0f&password=AutoCadIntegration67%3D&username=autocad_integration')
    console.log('JRBP -> resp:', resp);

  } catch (error) {
    console.log('JRBP -> error:', error);
  }
}