import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "0c472194-49de-476f-ad1f-8cd689bc60e9",
    authority: "https://login.microsoftonline.com/f285c292-6d27-443d-a1f0-dd70bcc4de16",
    redirectUri: "https://cordusbailey02.github.io/AddinTest/taskpane.html",
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

(async () => {
  await msalInstance.initialize();
  await msalInstance.handleRedirectPromise();
})();