import { appState } from "./appState";

import {
  getSharePointSiteId,
  getSharePointDriveId
} from "./sharePoint";

import {
  showMainSection,
  setStatus,
  setPaymentsStatus,
  showLoading,
  hideLoading,
  showNotification
} from "./uiHelpers";

import {
  browsers,
  browseFolder,
} from "./fileBrowser";

// Switch this when developing locally vs deploying
const IS_LOCAL    = true; // ← set to false before npm run build
const BASE_URL    = IS_LOCAL
  ? "https://localhost:3000"
  : "https://cordusbailey02.github.io/AddinTest";


function signIn() {
  setStatus("Opening sign in window...");

  openAuthDialog(async (token, userName) => {
    appState.accessToken = token;

    // Get SharePoint site and drive IDs after authentication
    const siteId = await getSharePointSiteId();

    appState.weeklyBondsId = await getSharePointDriveId(
      siteId,
      "Weekly Bonds"
    );

    appState.dailyLedgersId = await getSharePointDriveId(
      siteId,
      "Daily Ledgers"
    );

    showMainSection(userName);
    setStatus("");

    // Start both browsers at the ROOT of their respective libraries
    await browseFolder("submission", "root");
    await browseFolder("payments", "root");
  });
}

function signOut() {
  appState.accessToken              = null;
  appState.selectedFileId           = null;
  appState.selectedFileParentId     = null;
  appState.selectedFileName         = null;
  appState.selectedPaymentsFileId   = null;
  appState.selectedPaymentsFileParentId = null;
  appState.selectedPaymentsFileName = null;
  appState.pendingImportRows        = [];
  appState.pendingPaymentRows       = [];
  appState.submissionImportDone     = false;
  appState.paymentsImportDone       = false;
  appState.submissionSourceRows     = [];
  appState.ledgerSourceRows         = [];

  browsers.submission.stack = [{ id: "root", name: "My Files" }];
  browsers.payments.stack   = [{ id: "root", name: "My Files" }];

  document.getElementById("sign-in-section").style.display            = "block";
  document.getElementById("main-section").style.display               = "none";
  document.getElementById("preview-section").style.display            = "none";
  document.getElementById("preview-payments-section").style.display   = "none";
  document.getElementById("session-expired-banner").style.display     = "none";
  setStatus("");
  setPaymentsStatus("");
}

function openAuthDialog(onSuccess) {
  Office.context.ui.displayDialogAsync(
    `${BASE_URL}/dialog.html`,
    { height: 60, width: 30, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setStatus("Failed to open sign in dialog: " + result.error.message);
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (msg) => {
        dialog.close();
        try {
          const data = JSON.parse(msg.message);
          if (data.status === "success") {
            await onSuccess(data.accessToken, data.userName);
          } else {
            setStatus("Sign in error: " + data.message);
          }
        } catch (e) {
          setStatus("Error parsing auth response.");
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
        if (evt.error === 12006) setStatus("Sign in was cancelled.");
      });
    }
  );
}

async function refreshTokenViaDialog() {
  showLoading("Refreshing session...");
  openAuthDialog(async (token) => {
    appState.accessToken = token;
    hideLoading();
    showNotification("Session refreshed successfully.", "success");
    document.getElementById("session-expired-banner").style.display = "none";
  });
}

export {
  BASE_URL,
  signIn,
  signOut,
  refreshTokenViaDialog
};