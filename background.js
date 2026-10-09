// Background service worker for Recipient Manager
// Handles extension icon click to open the overlay panel

chrome.action.onClicked.addListener(async (tab) => {
  console.log("[RM Background] Icon clicked, tab:", tab);
  
  if (!tab || !tab.id) {
    console.log("[RM Background] No tab found");
    return;
  }

  const url = tab.url || "";
  console.log("[RM Background] Tab URL:", url);

  // First, try to inject the content script to ensure it's loaded
  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["content.js"]
    });
    console.log("[RM Background] Script injected/verified");
  } catch (err) {
    console.error("[RM Background] Failed to inject script:", err?.message || err);
    // Continue anyway - script might already be loaded
  }

  // Send message to content script
  console.log("[RM Background] Sending OPEN_RECIPIENT_MANAGER message to tab", tab.id);
  
  chrome.tabs.sendMessage(tab.id, { action: "OPEN_RECIPIENT_MANAGER" }, (response) => {
    if (chrome.runtime.lastError) {
      const errorMsg = chrome.runtime.lastError.message || JSON.stringify(chrome.runtime.lastError);
      console.error("[RM Background] Message error:", errorMsg);
      
      // If message failed, wait and try one more time
      setTimeout(() => {
        chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files: ["content.js"]
        }).then(() => {
          console.log("[RM Background] Script re-injected, retrying message...");
          setTimeout(() => {
            chrome.tabs.sendMessage(tab.id, { action: "OPEN_RECIPIENT_MANAGER" }, (response2) => {
              if (chrome.runtime.lastError) {
                const retryError = chrome.runtime.lastError.message || JSON.stringify(chrome.runtime.lastError);
                console.error("[RM Background] Retry failed:", retryError);
              } else {
                console.log("[RM Background] Retry response:", response2);
              }
            });
          }, 150);
        }).catch(err => {
          console.error("[RM Background] Inject failed:", err?.message || err);
        });
      }, 100);
    } else {
      console.log("[RM Background] Message response:", response);
    }
  });
});
