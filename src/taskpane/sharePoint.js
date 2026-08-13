import { appState } from "./appState";

async function graphFetch(path, driveId, options = {}) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/${path}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${appState.accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (response.status === 401) {
    document.getElementById("session-expired-banner").style.display = "block";
    throw new Error("Session expired. Please refresh your session using the button above.");
  }

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `Graph request failed: ${path}`);
  }

  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

async function getSharePointDriveId(siteId, libraryName) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    {
      headers: {
        Authorization: `Bearer ${appState.accessToken}`
      }
    }
  );

  if (!response.ok) {
    throw new Error(`Failed to get SharePoint drives: ${response.status}`);
  }

  const data = await response.json();

  //console.log("SharePoint drives:", data.value);

  const drive = data.value.find(
    d => d.name === libraryName
  );

  if (!drive) {
    throw new Error(`SharePoint document library "${libraryName}" not found`);
  }

  //console.log(`${libraryName} DRIVE ID:`, drive.id);

  return drive.id;
}

async function getSharePointSiteId() {
  const response = await fetch(
    "https://graph.microsoft.com/v1.0/sites/baileydevelopment.sharepoint.com:/sites/Bonds",
    {
      headers: {
        Authorization: `Bearer ${appState.accessToken}`
      }
    }
  );

  if (!response.ok) {
    throw new Error(`Failed to get SharePoint site: ${response.status}`);
  }

  const data = await response.json();

  //console.log("SHAREPOINT SITE ID:", data.id);

  return data.id;
}

export {
  graphFetch,
  getSharePointDriveId,
  getSharePointSiteId
};