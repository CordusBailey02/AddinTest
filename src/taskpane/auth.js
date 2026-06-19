import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "0c472194-49de-476f-ad1f-8cd689bc60e9",
    authority: "https://login.microsoftonline.com/f285c292-6d27-443d-a1f0-dd70bcc4de16",
    redirectUri: "https://cordusbailey02.github.io/AddinTest/auth.html",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: true,
  },
};

async function handleAuth() {
  const msalInstance = new PublicClientApplication(msalConfig);
  await msalInstance.initialize();
  
  // This processes the auth code in the URL and signals back to the opener
  const result = await msalInstance.handleRedirectPromise();
  
  if (result) {
    // Successfully processed - close the popup
    window.close();
  }
}

handleAuth().catch(console.error);