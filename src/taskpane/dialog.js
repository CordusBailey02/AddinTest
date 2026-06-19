import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "0c472194-49de-476f-ad1f-8cd689bc60e9",
    authority: "https://login.microsoftonline.com/f285c292-6d27-443d-a1f0-dd70bcc4de16",
    redirectUri: "https://cordusbailey02.github.io/AddinTest/dialog.html",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

const loginRequest = {
  scopes: ["User.Read", "Files.Read", "Files.Read.All", "Sites.Read.All"],
};

async function run() {
  const msalInstance = new PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Check if we're returning from a redirect
  const result = await msalInstance.handleRedirectPromise();

  if (result) {
    // We have a token — send it back to the taskpane and close
    Office.context.ui.messageParent(JSON.stringify({
      status: "success",
      accessToken: result.accessToken,
      userName: result.account.name || result.account.username,
    }));
  } else {
    // Not yet logged in — kick off the redirect login
    await msalInstance.loginRedirect(loginRequest);
  }
}

Office.onReady(() => {
  run().catch((error) => {
    Office.context.ui.messageParent(JSON.stringify({
      status: "error",
      message: error.message,
    }));
  });
});