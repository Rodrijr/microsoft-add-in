
let pca = undefined;
Office.onReady(async (info) => {
  if (info.host) {
    try {

      pca = new msal.PublicClientApplication({
        auth: {
          clientId: '637ff242-9444-4a30-9c55-e0b516887b4e',
          authority: 'https://login.microsoftonline.com/f3df3acc-6a4a-4618-adfd-8828f324887f',
        },
      });
    } catch (e) {
      console.log('JRBP -> e:', e);
    }
    run();
  }
});

async function run() {
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["Files.Read", "User.Read", "openid", "profile"],
    scopes: [],
  };
  let accessToken = null;

  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }

  if (accessToken === null) {
    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      accessToken = userAccount.accessToken;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
    }
  }

  // Call the Microsoft Graph API with the access token.
  const response = await fetch(
    `https://dev219430.service-now.com/login.do`,
    {
      headers: { Authorization: 'Bearer ' +accessToken },
    }
  );
  console.log('responseeeeeeeeeeeeeee', response)

}