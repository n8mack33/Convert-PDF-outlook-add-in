import { createNestablePublicClientApplication } from "@azure/msal-browser";

let pca = undefined;
Office.onReady(async (info) => {
  if (info.host) {
    // Initialize the public client application
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId:"a22056e3-1948-4b2f-a972-c0b806fc9cd1",//8c2474b9-ec3f-44d9-bc3f-1dbcdce29fd2 || process.env.Login_clint_id
        authority: `https://login.microsoftonline.com/common`,
        //  redirectUri: `${process.env.Local_RedirectBaseUrl}/assets/redirect.html`,
      },
    });
  }
});
// console.log(`redirectUri: ${process.env.Local_RedirectBaseUrl}/assets/redirect.html`);

 export async function Get_Token_SSO() {
    // Specify minimum scopes needed for the access token.
    const tokenRequest = {
      scopes: ["Files.ReadWrite", "User.Read", "profile", "openid","Mail.ReadWrite"],
    }; 
    let accessToken = null;
    
    // TODO 1: Call acquireTokenSilent.
    try {
        console.log("Trying to acquire token silently...");
        const userAccount = await pca.acquireTokenSilent(tokenRequest);
        console.log("Acquired token silently.");
        accessToken = userAccount.accessToken;
      } catch (error) {
        console.log(`Unable to acquire token silently: ${error}`);
      }
    // TODO 2: Call acquireTokenPopup.
    if (accessToken === null) {
        // Acquire token silent failure. Send an interactive request via popup.
        try {
          console.log("Trying to acquire token interactively...");
          const userAccount = await pca.acquireTokenPopup(tokenRequest);
          console.log("Acquired token interactively.");
          accessToken = userAccount.accessToken;
          console.log(userAccount);    
        } catch (popupError) {
          // Acquire token interactive failure.
          console.log(`Unable to acquire token interactively: ${popupError}`);
        }
      }
    // TODO 3: Log error if token still null.
    // Log error if both silent and popup requests failed.
if (accessToken === null) {
    console.error(`Unable to acquire access token.`);
    return;
  }
  return accessToken

    }
