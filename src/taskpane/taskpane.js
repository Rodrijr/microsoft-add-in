import { createNestablePublicClientApplication } from "@azure/msal-browser";

let pca = undefined;
Office.onReady(async (info) => {
  if (info.host) {
    /* document.getElementById("sideload-msg").style.display = "none";
     document.getElementById("app-body").style.display = "flex";
     document.getElementById("run").onclick = run;
 */
    // Initialize the public client application
    try {

      pca = await createNestablePublicClientApplication({
        auth: {
          clientId: "f5721a40-33b8-4b2b-8470-44db5b7813fa",
          authority: "https://login.microsoftonline.com/9dfb1a05-5f1d-449a-8960-62abcb479e7d"
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
    // scopes: ["Files.Read", "User.Read", "openid", "profile"],
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
    `https://iadbdev.service-now.com/login.do`,
    {
      headers: { Authorization: accessToken },
    }
  );
  console.log('responseeeeeeeeeeeeeee', response)
  /*if (response.ok) {
    // Write file names to the console.
    const data = await response.json();
    const names = data.value.map((item) => item.name);

    // Be sure the taskpane.html has an element with Id = item-subject.
    const label = document.getElementById("item-subject");

    // Write file names to task pane and the console.
    const nameText = names.join(", ");
    if (label) label.textContent = nameText;
    console.log(nameText);
  } else {
    const errorText = await response.text();
    console.error("Microsoft Graph call failed - error text: " + errorText);
  }*/

}