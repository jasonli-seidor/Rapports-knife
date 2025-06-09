const targetPage = "https://intranetnew.seidor.com/rapports/imputation-hours";

async function updateActionStatus(tabId) {
  try {
    const tab = await chrome.tabs.get(tabId);
    // The action should be enabled only on the target page.
    if (tab.url && tab.url.startsWith(targetPage)) {
      await chrome.action.enable(tabId);
    } else {
      await chrome.action.disable(tabId);
    }
  } catch (error) {
    // This can happen if the tab is closed before the check completes.
    console.log(`Could not update action for tab ${tabId}: ${error.message}`);
  }
}

// Update the action status when the user switches to a different tab.
chrome.tabs.onActivated.addListener(async (activeInfo) => {
  await updateActionStatus(activeInfo.tabId);
});

// Update the action status when a tab's URL changes.
chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  if (changeInfo.url) {
    await updateActionStatus(tabId);
  }
});

// When the extension is first installed, check all open tabs to set the initial state.
chrome.runtime.onInstalled.addListener(async () => {
  const tabs = await chrome.tabs.query({});
  for (const tab of tabs) {
    if (tab.id) {
      await updateActionStatus(tab.id);
    }
  }
});
