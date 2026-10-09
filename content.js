// ========= Content Script for Recipient Manager =========
// This script runs on Outlook Web pages and manages the overlay panel

// Guard against multiple injections
if (window.__rmContentScriptLoaded) {
  console.log("[RM Content] Already loaded, skipping re-injection");
} else {
  window.__rmContentScriptLoaded = true;

console.log("[RM Content] Script loaded");

// ========= Extraction helpers (Outlook Web) =========

// Get recipients from the Outlook web "well" (TO / CC)
function getRecipientsFromWell(type) {
  // Prefer stable ARIA labels; fall back to the older id suffix approach
  const wells = {
    TO: "div[aria-label='To']",
    CC: "div[aria-label='Cc']"
  };

  let well =
    document.querySelector(wells[type]) ||
    document.querySelector(`div[id$="_${type}"]`);

  // Additional fallback selectors for different Outlook layouts
  if (!well) {
    // Try finding by role and label
    well = document.querySelector(`div[role="group"][aria-label="${type === 'TO' ? 'To' : 'Cc'}"]`);
  }
  
  if (!well) {
    // Try finding the compose recipient container
    const composeArea = document.querySelector('div[aria-label="Message body"]')?.closest('div');
    if (composeArea) {
      const recipientContainer = composeArea.previousElementSibling;
      if (recipientContainer) {
        const label = type === 'TO' ? 'To' : 'Cc';
        const input = recipientContainer.querySelector(`input[aria-label*="${label}"], div[aria-label*="${label}"]`);
        if (input) well = input.closest('div');
      }
    }
  }

  if (!well) {
    console.log(`[RM Content] Could not find ${type} well`);
    return "";
  }

  console.log(`[RM Content] Found ${type} well:`, well);

  // Each recipient appears as span.lpcWrapper with an aria-label
  let wrappers = well.querySelectorAll("span.lpcWrapper[aria-label]");
  
  // Fallback: try other common recipient element patterns
  if (wrappers.length === 0) {
    wrappers = well.querySelectorAll("span[aria-label][title]");
  }
  if (wrappers.length === 0) {
    wrappers = well.querySelectorAll("button[aria-label]");
  }

  console.log(`[RM Content] Found ${wrappers.length} recipient wrappers in ${type}`);

  const recipients = Array.from(wrappers)
    .map(node => {
      let label = node.getAttribute("aria-label") || node.getAttribute("title") || "";

      // Remove leading icon / bullets / emoji-like artifacts Outlook adds
      // (e.g. , ❍, ○). Keep letters, numbers, and '(' as the start.
      label = label.replace(/^[^\w(]+/, "");

      // Normalize whitespace
      return label.trim();
    })
    .filter(Boolean);

  console.log(`[RM Content] Extracted recipients from ${type}:`, recipients);
  return recipients.join("; ");
}

// Extract greeting name from email body

function getGreetingNameFromBody() {
  const bodyNode = document.querySelector('div[role="document"]');
  if (!bodyNode) return "";

  const text = bodyNode.innerText || "";

  // Supports:
  //   Hi David,
  //   Hello Sarah,
  //   Hey Mike,
  //   Hello, David,
  //   Hey, John,
  //
  // Explanation:
  // - (hi|hello|hey)   = greeting word
  // - \b\s*,?\s+       = optional comma after greeting, then spaces
  // - ([A-Za-z][A-Za-z'\- ]{0,49}) = name (1–50 chars, letters + ' and - and spaces)
  // - \s*,             = comma after the name
  const match = text.match(/\b(hi|hello|hey)\b\s*,?\s+([A-Za-z][A-Za-z'\- ]{0,49})\s*,/i);

  if (match && match[2]) {
    return match[2].trim();
  }

  return "";
}

// Extract Client (folder) name from selected folder
function getClientFolderName() {
  const selected = document.querySelector('[role="treeitem"][aria-selected="true"]');
  if (!selected) return "";

  // 1) Prefer the visible folder text (original formatting)
  const labelSpan = selected.querySelector("span");
  if (labelSpan && labelSpan.innerText) {
    return labelSpan.innerText.trim();
  }

  // 2) Fallback: data-folder-name (may be lowercase)
  const folderKey = selected.getAttribute("data-folder-name");
  if (folderKey) {
    return folderKey.trim();
  }

  // 3) Last fallback: title before dash, e.g. "Ascension - 85 items (0 unread)"
  const title = selected.getAttribute("title");
  if (title) {
    return title.split("-")[0].trim();
  }

  return "";
}

// Get a display name for a folder node
function getFolderDisplayNameFromNode(node) {
  if (!node) return "";

  // Prefer visible span text
  const labelSpan = node.querySelector("span");
  if (labelSpan && labelSpan.innerText) {
    return labelSpan.innerText.trim();
  }

  // Fallback: title before dash
  const title = node.getAttribute("title");
  if (title) {
    return title.split("-")[0].trim();
  }

  // Last fallback: data-folder-name
  const folderKey = node.getAttribute("data-folder-name");
  if (folderKey) {
    return folderKey.trim();
  }

  return "";
}

// Normalize folder/client name for matching
function normalizeFolderName(name) {
  return String(name || "")
    .toLowerCase()
    // drop parentheses segments like " (A104)"
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// For matching only: strip short corporate prefixes like "DLP," from client list entries
function prepareClientNameForMatching(name) {
  let n = String(name || "");

  // Example: "DLP, Conemaugh Memorial Medical Center"
  const parts = n.split(",");
  if (parts.length > 1) {
    const first = parts[0].trim();

    // If the first segment is short (typical acronym: 2–4 chars), treat it as an owner prefix
    if (first.length > 0 && first.length <= 4) {
      // Drop the first segment and keep the rest
      n = parts.slice(1).join(",").trim();
    }
  }

  return n;
}

// Tokens that are very common in healthcare organization names and should not dominate the match
const RM_GENERIC_TOKENS = new Set([
  "inc",
  "llc",
  "corp",
  "corporation",
  "company",
  "co",
  "university",
  "univ",
  "hospital",
  "hosp",
  "medical",
  "med",
  "center",
  "ctr",
  "health",
  "system",
  "clinic",
  "group",
  "services",
  "service",

  // NEW: generic linker / stop words
  "the",
  "of",
  "and",
  "for",
  "at",
  "in"
]);

// Normalize common abbreviations so they compare better (e.g., "univ" ~ "university")
const RM_TOKEN_NORMALIZATION = {
  univ: "university",
  hosp: "hospital",
  med: "medical",
  ctr: "center",

  // NEW: plural forms → singular
  hospitals: "hospital",
  clinics: "clinic",
  systems: "system"
};

// Break a normalized folder/client name into tokens for fuzzy matching
function tokenizeName(name) {
  return String(name || "")
    .split(/[^a-z0-9]+/) // split on spaces, dashes, punctuation, etc.
    .map(part => part.trim())
    .filter(Boolean)
    .map(tok => RM_TOKEN_NORMALIZATION[tok] || tok); // normalize abbreviations
}

// Compute a fuzzy match score between target client name and a folder name
function scoreFolderMatch(targetNorm, folderNorm) {
  if (!targetNorm || !folderNorm) return 0;

  // Exact match: best possible score
  if (targetNorm === folderNorm) {
    return 100;
  }

  let score = 0;

  // Strong substring / "contains" relationship on the full normalized strings
  if (folderNorm.includes(targetNorm) || targetNorm.includes(folderNorm)) {
    score += 40;
  }

  const targetTokens = tokenizeName(targetNorm);
  const folderTokens = tokenizeName(folderNorm);

  if (targetTokens.length && folderTokens.length) {
    const targetSet = new Set(targetTokens);
    let intersection = 0;

    folderTokens.forEach(tok => {
      if (targetSet.has(tok)) intersection++;
    });

    const unionSize = new Set([...targetTokens, ...folderTokens]).size;
    const jaccard = unionSize ? intersection / unionSize : 0; // 0–1

    // Up to 40 points based on word overlap
    score += Math.round(jaccard * 40);

    // Bonus if first tokens match exactly (common strongest signal)
    if (targetTokens[0] && folderTokens[0] && targetTokens[0] === folderTokens[0]) {
      score += 10;
    }

    // --- Brand token logic: favor the main "brand" word like "rush", "aspirus" ---
    const targetNonGeneric = targetTokens.filter(t => !RM_GENERIC_TOKENS.has(t));
    const folderNonGeneric = folderTokens.filter(t => !RM_GENERIC_TOKENS.has(t));

    if (targetNonGeneric.length > 0) {
      const mainBrand = targetNonGeneric[0]; // e.g., "rush" in "rush university medical center"

      let brandFound = false;
      let brandAtFront = false;

      folderNonGeneric.forEach((tok, idx) => {
        const matchesBrand =
          tok === mainBrand ||
          tok.startsWith(mainBrand) ||
          mainBrand.startsWith(tok);

        if (matchesBrand) {
          brandFound = true;
          if (idx === 0) {
            brandAtFront = true; // brand is first non-generic word in folder name
          }
        }
      });

      if (brandFound) {
        // Stronger bonus if the brand is the FIRST non-generic token,
        // e.g., "Valley Health" vs "Mohawk Valley Health"
        score += brandAtFront ? 35 : 15;
      } else {
        // Slight penalty if it does not, so "generic-only" folders rank lower
        score -= 10;
      }
    }
  }

  // Never return negative scores
  return Math.max(0, score);
}

// Minimum fuzzy score for a "strong" match
const RM_MIN_CLIENT_MATCH_SCORE = 25;

// Minimum score to accept a "closest match" at all
const RM_MIN_CLOSEST_MATCH_SCORE = 10;

// Find the best matching folder node for a given client name (fuzzy match + closest match)
function findFolderNodeForClient(clientName) {
  rmLastClientFolderMatch = null; // reset each time

  if (!clientName) return null;

  // FIX: drop short corporate prefixes ("DLP,") before we normalize & match
  const clientForMatching = prepareClientNameForMatching(clientName);
  const target = normalizeFolderName(clientForMatching);
  if (!target) return null;

  const nodes = document.querySelectorAll('[role="treeitem"]');
  let bestNode = null;
  let bestScore = 0;

  nodes.forEach(node => {
    const display = getFolderDisplayNameFromNode(node);
    if (!display) return;

    const norm = normalizeFolderName(display);
    if (!norm) return;

    const score = scoreFolderMatch(target, norm);

    if (score > bestScore) {
      bestScore = score;
      bestNode = node;
    }
  });

  // No candidate at all
  if (!bestNode || bestScore <= 0) {
    return null;
  }

  // If even the best score is below the "closest match" floor, treat as no match
  if (bestScore < RM_MIN_CLOSEST_MATCH_SCORE) {
    return null;
  }

  // Record info about this best match so the UI can choose the right status message
  const displayName = getFolderDisplayNameFromNode(bestNode) || clientName;

  rmLastClientFolderMatch = {
    clientNameInput: clientName,
    folderDisplayName: displayName,
    score: bestScore,
    isWeakMatch: bestScore < RM_MIN_CLIENT_MATCH_SCORE
  };

  return bestNode;
}

// Jump to folder for client name; returns the node if found, else null
function jumpToClientFolder(clientName) {
  const node = findFolderNodeForClient(clientName);
  if (!node) return null;

  try {
    node.scrollIntoView({ block: "center", behavior: "smooth" });
  } catch (e) {
    // older browsers may not support smooth
    node.scrollIntoView({ block: "center" });
  }
  node.click();
  return node;
}

// Main extractor
function extractOutlookEmailInfo() {
  const email = getRecipientsFromWell("TO");
  const cc = getRecipientsFromWell("CC");
  const name = getGreetingNameFromBody();
  const client = getClientFolderName();

  return {
    clientName: client,
    greetingName: name,
    email: email,
    cc: cc
  };
}

// ========= Panel state =========

let rmPanelContainer = null;
let rmEntries = [];
let rmSessionActive = false;
let rmLastDeleted = null; // { entry, index } for Undo
let rmLastCleared = null;   // stores entire list for Undo after Clear List
let rmPanelMinimized = false; // minimize state

// NEW: client navigation state
let rmClientSequence = [];      // array of client names
let rmClientSequenceIndex = -1; // current index in sequence (-1 = not started)

// NEW: info about the last folder match
let rmLastClientFolderMatch = null; 
// { clientNameInput, folderDisplayName, score, isWeakMatch }

// ========= Message from background: open panel =========

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log("[RM Content] Message received:", message);
  
  if (message.action === "OPEN_RECIPIENT_MANAGER") {
    console.log("[RM Content] Opening panel...");
    try {
      openRecipientManagerPanel();
      sendResponse({ success: true });
    } catch (err) {
      console.error("[RM Content] Error opening panel:", err);
      sendResponse({ success: false, error: err.message });
    }
  }
  return true; // Keep message channel open for async response
});

// ========= Panel creation =========

function openRecipientManagerPanel() {
  console.log("[RM Content] openRecipientManagerPanel called");
  
  if (rmPanelContainer && document.body.contains(rmPanelContainer)) {
    console.log("[RM Content] Panel already exists, showing it");
    rmPanelContainer.style.display = "block";
    return;
  }

  console.log("[RM Content] Creating new panel...");
  
  // Inject CSS if not already present
  injectRmStyles();

  // Create container
  rmPanelContainer = document.createElement("div");
  rmPanelContainer.id = "rm-panel-container";

  const panel = document.createElement("div");
  panel.className = "rm-panel";

  panel.innerHTML = `
    <div class="rm-panel__header">
      <h2 class="rm-panel__title">Recipeint Manager - Builds Email Lists</h2>
      <div class="rm-header-controls">
        <button id="rm-minimizeButton" class="rm-panel__minimize" aria-label="Minimize panel" title="Minimize">&#8722;</button>
        <button id="rm-closeButton" class="rm-panel__close" aria-label="Close Recipient Manager">&times;</button>
      </div>
    </div>
    <div class="rm-panel__body">
      <div class="rm-row">
        <label for="rm-clientName">Client:</label>
        <input id="rm-clientName" type="text" placeholder="Start typing a client..." />
      </div>

      <div class="rm-row">
        <label for="rm-nameField">Name:</label>
        <input id="rm-nameField" type="text" placeholder="Contact name" />
      </div>

      <div class="rm-row">
        <label for="rm-emailField">Email:</label>
        <textarea id="rm-emailField" rows="2" placeholder="To addresses; use ; to separate"></textarea>
      </div>

      <div class="rm-row">
        <label for="rm-ccField">Cc:</label>
        <textarea id="rm-ccField" rows="2" placeholder="Cc addresses; use ; to separate"></textarea>
      </div>

      <div class="rm-button-row">
        <div class="rm-left-buttons">
          <button id="rm-btnExtract" class="rm-btn-primary" type="button">
            <span class="btn-ic" aria-hidden="true">
              <svg viewBox="0 0 24 24"><path d="M12 3a1 1 0 0 1 1 1v9.586l2.293-2.293a1 1 0 1 1 1.414 1.414l-4.007 4.007a1 1 0 0 1-1.414 0L7.28 12.707a1 1 0 1 1 1.414-1.414L11 13.586V4a1 1 0 0 1 1-1zM5 19a1 1 0 1 0 0 2h14a1 1 0 1 0 0-2H5z"/></svg>
            </span>
            Extract
          </button>
          <button id="rm-btnClear" class="rm-btn-primary" type="button">
            <span class="btn-ic" aria-hidden="true">
              <svg viewBox="0 0 24 24"><path d="M9 3h6a2 2 0 0 1 2 2v1h3a1 1 0 1 1 0 2h-1.08l-1.13 11.3A2 2 0 0 1 15.8 22H8.2a2 2 0 0 1-1.99-1.7L5.08 8H4a1 1 0 1 1 0-2h3V5a2 2 0 0 1 2-2zm6 3V5H9v1h6zm-5.5 5a1 1 0 0 0-1 1l.5 6a1 1 0 0 0 2 0l-.5-6a1 1 0 0 0-1-1zm5 0a1 1 0 0 0-1 1l-.5 6a1 1 0 0 0 2 0l.5-6a1 1 0 0 0-1-1z"/></svg>
            </span>
            Clear
          </button>
        </div>

        <div class="rm-left-buttons rm-client-config">
          <button id="rm-btnConfigureClientList" class="rm-btn-primary" type="button">
            <span class="btn-ic" aria-hidden="true">
              <svg viewBox="0 0 24 24"><path d="M4 6h2v2H4V6zm4 1h12v1H8V7zM4 11h2v2H4v-2zm4 1h12v1H8v-1zM4 16h2v2H4v-2zm4 1h12v1H8v-1z"/></svg>
            </span>
            Client List...
          </button>
          <div class="rm-client-nav-row">
            <button id="rm-btnPrev" class="rm-btn-secondary" type="button" disabled>
              <span class="btn-ic" aria-hidden="true">
                <svg viewBox="0 0 24 24"><path d="M15.41 7.41 14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>
              </span>
              Prev
            </button>
            <button id="rm-btnNext" class="rm-btn-secondary" type="button" disabled>
              <span class="btn-ic" aria-hidden="true">
                <svg viewBox="0 0 24 24"><path d="m8.59 16.59 1.41 1.41 6-6-6-6-1.41 1.41L13.17 12z"/></svg>
              </span>
              Next
            </button>
          </div>
        </div>

        <div class="rm-right-buttons">
          <button id="rm-btnSendToList" class="rm-btn-primary" type="button">
            <span class="btn-ic" aria-hidden="true">
              <svg viewBox="0 0 24 24"><path d="M4 5h12v2H4zM4 11h12v2H4zM4 17h12v2H4zM20 5h-2v14h2z"/></svg>
            </span>
            <span class="btn-text">Build List Off</span>
          </button>
          <div class="rm-right-buttons-bottom">
            <button id="rm-btnAdd" class="rm-btn-secondary" type="button" disabled title="Add">+Add</button>
            <button id="rm-btnClearList" class="rm-btn-secondary" type="button" title="Clear List">Clear List</button>
          </div>
        </div>
      </div>

      <div id="rm-status" class="info">Ready.</div>

      <h3 class="rm-list-heading">Current List</h3>
      <div id="rm-tableContainer" class="tbl__wrap tbl--responsive">
        <!-- table injected here -->
      </div>

      <div class="rm-export-buttons">
        <button id="rm-btnCopyTsv" class="rm-btn-primary" type="button">
          <span class="btn-ic" aria-hidden="true">
            <svg viewBox="0 0 24 24"><path d="M16 1H4a2 2 0 0 0-2 2v12h2V3h12V1zm3 4H8a2 2 0 0 0-2 2v14h13a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2zm0 16H8V7h11v14z"/></svg>
          </span>
          Copy
        </button>
        <label><input type="checkbox" id="rm-includeHeaders" checked /> Include headers</label>
        <label><input type="checkbox" id="rm-includeClient" checked /> Include client</label>
        <button id="rm-btnUndo" class="rm-btn-secondary" type="button" disabled>
          <span class="btn-ic" aria-hidden="true">
            <svg viewBox="0 0 24 24"><path d="M12 5V1L5 8l7 7v-4c5.523 0 9 1.477 11 5-1-7-5.477-11-11-11z"/></svg>
          </span>
          Undo
        </button>
      </div>

      <div id="rm-clientListOverlay" class="rm-clientListOverlay" style="display:none;">
        <div class="rm-clientListCard">
          <h3 class="rm-clientListTitle">Client List</h3>
          <p class="rm-clientListHelp">Paste or type client names, one per line (up to 100).</p>
          <textarea id="rm-clientListTextarea" rows="12" placeholder=""></textarea>
          <div class="rm-clientListButtons">
            <button id="rm-clientListDone" class="rm-btn-primary" type="button">
              <span class="btn-ic" aria-hidden="true">
                <svg viewBox="0 0 24 24"><path d="M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"/></svg>
              </span>
              Done
            </button>
            <button id="rm-clientListClear" class="rm-btn-secondary" type="button">
              <span class="btn-ic" aria-hidden="true">
                <svg viewBox="0 0 24 24"><path d="M16.24 3.56 21 8.32 10.83 18.5H6.07L1.31 13.74 12.8 2.24a2 2 0 0 1 2.83 0l.61.61zM7.49 16.5h2.34l8.67-8.67-2.34-2.34L5.15 14.16l2.34 2.34zM22 20.5H10.83l2-2H22v2z"/></svg>
              </span>
              Clear
            </button>
            <button id="rm-clientListCancel" class="rm-btn-secondary" type="button">
              <span class="btn-ic" aria-hidden="true">
                <svg viewBox="0 0 24 24"><path d="M18.3 5.71 12 12l6.3 6.29-1.41 1.42L10.59 13.41 4.29 19.71 2.88 18.3 9.17 12 2.88 5.71 4.29 4.29 10.59 10.59 16.89 4.29z"/></svg>
              </span>
              Cancel
            </button>
          </div>
        </div>
      </div>
    </div>
  `;

  rmPanelContainer.appendChild(panel);
  document.body.appendChild(rmPanelContainer);

  console.log("[RM Content] Panel added to DOM");

  wirePanelEvents(panel);
  renderRmTable();
}

// Inject extension stylesheet into page
function injectRmStyles() {
  if (document.getElementById("rm-styles")) {
    console.log("[RM Content] Styles already injected");
    return;
  }

  try {
    const link = document.createElement("link");
    link.id = "rm-styles";
    link.rel = "stylesheet";
    link.type = "text/css";
    link.href = chrome.runtime.getURL("styles.css");
    document.head.appendChild(link);
    console.log("[RM Content] Styles injected successfully");
  } catch (err) {
    console.error("[RM Content] Failed to inject styles:", err);
  }
}

// ========= Panel event wiring =========

function wirePanelEvents(panel) {
  const header = panel.querySelector(".rm-panel__header");
  const btnClose = panel.querySelector(".rm-panel__close");
  const btnMinimize = panel.querySelector("#rm-minimizeButton");
  const clientInput = panel.querySelector("#rm-clientName");
  const nameInput = panel.querySelector("#rm-nameField");
  const emailField = panel.querySelector("#rm-emailField");
  const ccField = panel.querySelector("#rm-ccField");
  const btnExtract = panel.querySelector("#rm-btnExtract");
  const btnClear = panel.querySelector("#rm-btnClear");
  const btnSendToList = panel.querySelector("#rm-btnSendToList");
  const btnAdd = panel.querySelector("#rm-btnAdd");
  const btnClearList = panel.querySelector("#rm-btnClearList");
  const btnCopyTsv = panel.querySelector("#rm-btnCopyTsv");
  const includeHeadersCheckbox = panel.querySelector("#rm-includeHeaders");
  const includeClientCheckbox = panel.querySelector("#rm-includeClient");
  const btnUndo = panel.querySelector("#rm-btnUndo");

  // NEW: client navigation + config
  const btnPrev = panel.querySelector("#rm-btnPrev");
  const btnNext = panel.querySelector("#rm-btnNext");
  const btnConfigureClientList = panel.querySelector("#rm-btnConfigureClientList");
  const clientListOverlay = panel.querySelector("#rm-clientListOverlay");
  const clientListTextarea = panel.querySelector("#rm-clientListTextarea");
  const btnClientListDone = panel.querySelector("#rm-clientListDone");
  const btnClientListClear = panel.querySelector("#rm-clientListClear");
  const btnClientListCancel = panel.querySelector("#rm-clientListCancel");

  // Initialize Prev/Next
  btnPrev.disabled = true;
  btnNext.disabled = true;

  const statusDiv = panel.querySelector("#rm-status");
  const tableContainer = panel.querySelector("#rm-tableContainer");

  function setStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = type || "info";
  }

  // --- Minimize panel ---
  btnMinimize.addEventListener("click", () => {
    rmPanelMinimized = !rmPanelMinimized;
    if (rmPanelMinimized) {
      panel.classList.add("minimized");
      btnMinimize.innerHTML = "+";
      btnMinimize.title = "Restore";
    } else {
      panel.classList.remove("minimized");
      btnMinimize.innerHTML = "&#8722;";
      btnMinimize.title = "Minimize";
    }
  });

  // Helper to strip trailing (something) from client/folder names
  function cleanClientName(rawName) {
    if (!rawName) return rawName;
    // Remove a trailing " ( ... )" at the end of the string
    return rawName.replace(/\s*\([^)]*\)\s*$/, "").trim();
  }

  // --- Client navigation helper functions ---
  function updateClientNavUi() {
    if (!rmClientSequence.length) {
      btnPrev.disabled = true;
      btnNext.disabled = true;
      return;
    }

    // disable Prev on first or not-started
    btnPrev.disabled = rmClientSequenceIndex <= 0;

    // If not started, Next should be enabled
    if (rmClientSequenceIndex === -1) {
      btnNext.disabled = false;
    } else {
      btnNext.disabled = rmClientSequenceIndex >= rmClientSequence.length - 1;
    }
  }

  function openClientListOverlay() {
    if (!clientListOverlay) return;

    // Pre-fill with existing sequence
    if (clientListTextarea) {
      clientListTextarea.value = rmClientSequence.join("\n");
    }

    clientListOverlay.style.display = "flex";
  }

  function closeClientListOverlay() {
    if (!clientListOverlay) return;
    clientListOverlay.style.display = "none";
  }

  // --- Close panel ---
  btnClose.addEventListener("click", () => {
    if (rmPanelContainer && rmPanelContainer.parentNode) {
      rmPanelContainer.parentNode.removeChild(rmPanelContainer);
      rmPanelContainer = null;
    }
  });

  // --- Make panel draggable via header ---
  let isDragging = false;
  let dragOffsetX = 0;
  let dragOffsetY = 0;

  function onMouseMove(e) {
    if (!isDragging || !rmPanelContainer) return;
    
    const panelRect = rmPanelContainer.querySelector('.rm-panel').getBoundingClientRect();
    const panelWidth = panelRect.width;
    const panelHeight = panelRect.height;
    
    // Calculate new position
    let newLeft = e.clientX - dragOffsetX;
    let newTop = e.clientY - dragOffsetY;
    
    // Constrain to viewport
    newLeft = Math.max(0, Math.min(newLeft, window.innerWidth - panelWidth));
    newTop = Math.max(0, Math.min(newTop, window.innerHeight - panelHeight));
    
    rmPanelContainer.style.left = newLeft + "px";
    rmPanelContainer.style.top = newTop + "px";
  }

  function onMouseUp() {
    if (!isDragging) return;
    isDragging = false;
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  }

  header.addEventListener("mousedown", (e) => {
    // Don't start drag if clicking the close button
    if (e.target.closest(".rm-panel__close")) return;

    if (!rmPanelContainer) return;
    const rect = rmPanelContainer.getBoundingClientRect();

    isDragging = true;
    dragOffsetX = e.clientX - rect.left;
    dragOffsetY = e.clientY - rect.top;

    // Switch from right-based to left-based positioning for drag
    rmPanelContainer.style.right = "auto";
    rmPanelContainer.style.left = rect.left + "px";
    rmPanelContainer.style.top = rect.top + "px";

    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  });

  // --- Extract from Outlook ---
  btnExtract.addEventListener("click", () => {
    setStatus("Extracting from Outlook web...", "info");
    try {
      const data = extractOutlookEmailInfo();

      // Use existing value as fallback, then clean trailing (...) from the folder name
      const rawClientName = (data.clientName || clientInput.value || "").trim();
      const cleanedClientName = cleanClientName(rawClientName);
      clientInput.value = cleanedClientName;

      nameInput.value = data.greetingName || nameInput.value || "";
      
      // Always set Email and Cc from extraction (To -> Email, Cc -> Cc)
      emailField.value = data.email || "";
      ccField.value = data.cc || "";

      // Show what was extracted in status
      const extracted = [];
      if (data.email) extracted.push(`To: ${data.email}`);
      if (data.cc) extracted.push(`Cc: ${data.cc}`);
      if (extracted.length > 0) {
        setStatus("Extracted: " + extracted.join("; "), "success");
      } else {
        setStatus("Extracted successfully (no To/Cc recipients found).", "info");
      }
    } catch (err) {
      setStatus("Error during extraction: " + (err.message || err), "error");
    }
  });

  // --- Clear form fields (not list) ---
  btnClear.addEventListener("click", () => {
    clientInput.value = "";
    nameInput.value = "";
    emailField.value = "";
    ccField.value = "";
    setStatus("Fields cleared.", "info");
  });

  // --- Toggle Send to List session ---
  const btnTextSpan = btnSendToList.querySelector(".btn-text");

  btnSendToList.addEventListener("click", () => {
    rmSessionActive = !rmSessionActive;

    if (rmSessionActive) {
      if (btnTextSpan) btnTextSpan.textContent = "Build List On";
      btnAdd.disabled = false;
      setStatus("Build List is ON. Use Add to append rows.", "info");
    } else {
      if (btnTextSpan) btnTextSpan.textContent = "Build List Off";
      btnAdd.disabled = true;
      setStatus("Build List is OFF. List remains visible.", "info");
    }
  });

  // --- Add current entry ---
  btnAdd.addEventListener("click", () => {
    if (!rmSessionActive) {
      setStatus("Turn ON 'Send to List' before adding.", "error");
      return;
    }

    const entry = {
      Client: (clientInput.value || "").trim(),
      Name: (nameInput.value || "").trim(),
      Email: (emailField.value || "").trim(),
      Cc: (ccField.value || "").trim()
    };

    if (!entry.Email && !entry.Cc) {
      setStatus("Nothing to add. Fill Email or Cc first.", "error");
      return;
    }

    const isDuplicate = rmEntries.some(e =>
      e.Client === entry.Client &&
      e.Name === entry.Name &&
      e.Email === entry.Email &&
      e.Cc === entry.Cc
    );

    if (isDuplicate) {
      setStatus("This entry is already in the list.", "info");
      return;
    }

    rmEntries.push(entry);
    renderRmTable();
    setStatus("Entry added to list.", "success");
  });

  // --- Clear list ---
  btnClearList.addEventListener("click", () => {
    if (rmEntries.length === 0) {
      setStatus("List is already empty.", "info");
      return;
    }

    const ok = window.confirm("Clear the entire list? This can be undone.");
    if (!ok) {
      setStatus("Clear List canceled.", "info");
      return;
    }

    // Save entire list for Undo
    rmLastCleared = [...rmEntries];
    rmLastDeleted = null;            // this is not a row delete

    rmEntries = [];
    renderRmTable();

    if (rmPanelContainer && rmPanelContainer.btnUndo) {
      rmPanelContainer.btnUndo.disabled = false;
    }

    setStatus("List cleared. Click Undo to restore.", "info");
  });

  // --- Copy TSV for Excel ---
  btnCopyTsv.addEventListener("click", async () => {
    if (rmEntries.length === 0) {
      setStatus("Nothing to copy. List is empty.", "error");
      return;
    }

    const includeHeaders = includeHeadersCheckbox && includeHeadersCheckbox.checked;
    const includeClient = includeClientCheckbox && includeClientCheckbox.checked;

    const lines = [];

    // Optional header row
    if (includeHeaders) {
      const header = [];
      if (includeClient) header.push("Client");
      header.push("Name", "Email", "Cc");
      lines.push(header.join("\t"));
    }

    // Data rows
    rmEntries.forEach(e => {
      const row = [];
      if (includeClient) row.push(e.Client || "");
      row.push(e.Name || "");
      row.push(e.Email || "");
      row.push(e.Cc || "");
      lines.push(row.join("\t"));
    });

    const tsv = lines.join("\n");

    try {
      await navigator.clipboard.writeText(tsv);

      const msg =
        includeHeaders && includeClient
          ? "Copied with headers and client column."
          : includeHeaders && !includeClient
          ? "Copied with headers but without client."
          : !includeHeaders && includeClient
          ? "Copied without headers but with client column."
          : "Copied rows only (no headers, no client).";

      setStatus(msg, "success");
    } catch (err) {
      setStatus("Could not copy to clipboard: " + (err.message || err), "error");
    }
  });

  // --- Undo last delete or Clear List ---
  btnUndo.addEventListener("click", () => {

    // --- Undo row delete ---
    if (rmLastDeleted) {
      const { entry, index } = rmLastDeleted;

      if (index >= 0 && index <= rmEntries.length) {
        rmEntries.splice(index, 0, entry);
      } else {
        rmEntries.push(entry);
      }

      rmLastDeleted = null;
      setStatus("Last removed row restored.", "info");
      renderRmTable();
    }

    // --- Undo Clear List ---
    else if (rmLastCleared) {
      rmEntries = [...rmLastCleared];
      rmLastCleared = null;

      setStatus("Cleared list restored.", "info");
      renderRmTable();
    }

    // Disable Undo if nothing left to undo
    if (!rmLastDeleted && !rmLastCleared) {
      btnUndo.disabled = true;
    }
  });

  // --- Open client list overlay ---
  btnConfigureClientList.addEventListener("click", () => {
    openClientListOverlay();
  });

  // --- Save client list and close overlay ---
  btnClientListDone.addEventListener("click", () => {
    if (!clientListTextarea) {
      closeClientListOverlay();
      return;
    }

    const raw = clientListTextarea.value || "";
    let names = raw.split(/\r?\n/)
      .map(line => line.trim())
      .filter(line => line);

    if (names.length > 100) {
      names = names.slice(0, 100);
    }

    rmClientSequence = names;
    rmClientSequenceIndex = -1;

    closeClientListOverlay();
    updateClientNavUi();

    if (names.length) {
      setStatus(`Client list loaded with ${names.length} client(s). Use Prev/Next to jump between folders.`, "success");
    } else {
      setStatus("Client list is empty. Prev/Next are disabled.", "info");
    }
  });

  // --- Clear client list textarea (does not close overlay) ---
  btnClientListClear.addEventListener("click", () => {
    if (clientListTextarea) {
      clientListTextarea.value = "";
    }
    setStatus("Client list cleared in editor.", "info");
  });

  // --- Cancel client list editing ---
  btnClientListCancel.addEventListener("click", () => {
    closeClientListOverlay();
    setStatus("Client list edit canceled.", "info");
  });

  // --- Navigate to previous client folder ---
  btnPrev.addEventListener("click", () => {
    if (!rmClientSequence.length) {
      setStatus("No client list configured. Click 'Client List...' to add clients.", "error");
      return;
    }

    if (rmClientSequenceIndex <= 0) {
      setStatus("Already at the first client in the list.", "info");
      return;
    }

    rmClientSequenceIndex--;
    const clientName = rmClientSequence[rmClientSequenceIndex];

    const node = jumpToClientFolder(clientName);
    if (node) {
      const displayName = getFolderDisplayNameFromNode(node) || clientName;
      const cleaned = cleanClientName(displayName);
      clientInput.value = cleaned;

      // Clear previously extracted/person-specific fields
      nameInput.value = "";
      emailField.value = "";
      ccField.value = "";

      if (rmLastClientFolderMatch && rmLastClientFolderMatch.isWeakMatch) {
        setStatus(`Closest match found: ${cleaned}`, "info");
      } else {
        setStatus(`Jumped to client folder: ${cleaned}`, "success");
      }
    } else {
      setStatus(`No folder found matching '${clientName}'.`, "error");
    }

    updateClientNavUi();
  });

  // --- Navigate to next client folder ---
  btnNext.addEventListener("click", () => {
    if (!rmClientSequence.length) {
      setStatus("No client list configured. Click 'Client List...' to add clients.", "error");
      return;
    }

    // If not started, move to first; else move forward
    let nextIndex = rmClientSequenceIndex < 0 ? 0 : rmClientSequenceIndex + 1;

    if (nextIndex >= rmClientSequence.length) {
      setStatus("Already at the last client in the list.", "info");
      updateClientNavUi();
      return;
    }

    rmClientSequenceIndex = nextIndex;
    const clientName = rmClientSequence[rmClientSequenceIndex];

    const node = jumpToClientFolder(clientName);
    if (node) {
      const displayName = getFolderDisplayNameFromNode(node) || clientName;
      const cleaned = cleanClientName(displayName);
      clientInput.value = cleaned;

      // Clear previously extracted/person-specific fields
      nameInput.value = "";
      emailField.value = "";
      ccField.value = "";

      if (rmLastClientFolderMatch && rmLastClientFolderMatch.isWeakMatch) {
        setStatus(`Closest match found: ${cleaned}`, "info");
      } else {
        setStatus(`Jumped to client folder: ${cleaned}`, "success");
      }
    } else {
      setStatus(`No folder found matching '${clientName}'.`, "error");
    }

    updateClientNavUi();
  });

  // Save references on container for table render
  rmPanelContainer.tableContainer = tableContainer;
  rmPanelContainer.statusDiv = statusDiv;
  rmPanelContainer.btnUndo = btnUndo;
}

// ========= Table rendering & row delete =========

function renderRmTable() {
  if (!rmPanelContainer || !rmPanelContainer.tableContainer) return;

  const container = rmPanelContainer.tableContainer;

  if (rmEntries.length === 0) {
    container.innerHTML = `<p style="font-size:11px;color:#666;">No entries yet. Turn on 'Build List' and click Add.</p>`;
    return;
  }

  const html = buildTableHtml(rmEntries, { includeWrapper: true, includeDeleteColumn: true });
  container.innerHTML = html;

  const table = container.querySelector("table.tbl");
  if (table) {
    initSortableTable(table);
    wireDeleteButtons(table);
  }
}

function buildTableHtml(data, options) {
  const includeWrapper = options && options.includeWrapper;
  const includeDelete = options && options.includeDeleteColumn;

  const rowsHtml = data.map((entry, index) => {
    const deleteCell = includeDelete
      ? `<td class="rm-actions-cell"><button type="button" class="rm-delete-btn" data-index="${index}">✕</button></td>`
      : "";

    return `
      <tr>
        ${deleteCell}
        <td>${escapeHtml(entry.Client || "")}</td>
        <td>${escapeHtml(entry.Name || "")}</td>
        <td>${escapeHtml(entry.Email || "")}</td>
        <td>${escapeHtml(entry.Cc || "")}</td>
      </tr>`;
  }).join("");

  // Empty header for delete column – no text
  const deleteHeader = includeDelete
    ? `<th class="rm-actions-header"></th>`
    : "";

  const tableCore = `
<table class="tbl tbl--striped tbl--bordered tbl--compact" aria-describedby="recipient-table-caption">
  <caption id="recipient-table-caption">Captured recipient data for the current session</caption>
  <thead>
    <tr>
      ${deleteHeader}
      <th scope="col" aria-sort="none" tabindex="0">Client</th>
      <th scope="col" aria-sort="none" tabindex="0">Name</th>
      <th scope="col" aria-sort="none" tabindex="0">Email</th>
      <th scope="col" aria-sort="none" tabindex="0">Cc</th>
    </tr>
  </thead>
  <tbody>
    ${rowsHtml}
  </tbody>
</table>`.trim();

  if (!includeWrapper) {
    return tableCore;
  }

  return `
<figure class="tbl__wrap tbl--responsive">
  ${tableCore}
</figure>`.trim();
}

function wireDeleteButtons(table) {
  const buttons = table.querySelectorAll(".rm-delete-btn");
  buttons.forEach(btn => {
    btn.addEventListener("click", () => {
      const indexStr = btn.getAttribute("data-index");
      const index = indexStr ? parseInt(indexStr, 10) : -1;
      if (index < 0 || index >= rmEntries.length) return;

      const ok = window.confirm("Remove this row from the list?");
      if (!ok) return;

      // Save last deleted for Undo
      rmLastDeleted = {
        entry: rmEntries[index],
        index: index
      };

      rmEntries.splice(index, 1);
      renderRmTable();

      // Enable Undo button
      if (rmPanelContainer && rmPanelContainer.btnUndo) {
        rmPanelContainer.btnUndo.disabled = false;
      }

      // Show notification
      if (rmPanelContainer && rmPanelContainer.statusDiv) {
        rmPanelContainer.statusDiv.textContent = "Row removed. Click Undo to restore.";
        rmPanelContainer.statusDiv.className = "info";
      }
    });
  });
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ========= Sorting logic =========

function initSortableTable(table) {
  const headerCells = table.querySelectorAll("thead th[scope='col']");

  headerCells.forEach((th, columnIndex) => {
    if (!th.hasAttribute("tabindex")) {
      th.setAttribute("tabindex", "0");
    }

    th.addEventListener("click", () => handleSort(table, th, columnIndex));
    th.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        handleSort(table, th, columnIndex);
      }
    });
  });
}

function handleSort(table, th, columnIndex) {
  const currentSort = th.getAttribute("aria-sort") || "none";
  const newSort = currentSort === "ascending" ? "descending" : "ascending";

  const allHeaders = table.querySelectorAll("thead th[scope='col']");
  allHeaders.forEach(header => {
    if (header !== th) {
      header.setAttribute("aria-sort", "none");
    }
  });

  th.setAttribute("aria-sort", newSort);
  sortTableByColumn(table, columnIndex, newSort === "ascending");
}

function sortTableByColumn(table, columnIndex, ascending) {
  const tbody = table.querySelector("tbody");
  if (!tbody) return;

  const rows = Array.from(tbody.querySelectorAll("tr"));

  // If there's a delete/actions column, the actual cell index is shifted by +1
  const hasDeleteColumn = !!table.querySelector("th.rm-actions-header");
  const cellIndex = hasDeleteColumn ? columnIndex + 1 : columnIndex;

  const rowData = rows.map((row, originalIndex) => {
    const cell = row.children[cellIndex];
    const raw =
      cell?.getAttribute("data-sort-value") ??
      (cell?.textContent || "").trim();

    return {
      row,
      originalIndex,
      value: raw
    };
  });

  const isNumeric = detectNumericColumn(rowData, table, columnIndex);

  rowData.sort((a, b) => {
    let cmp;

    if (isNumeric) {
      const aNum = parseFloat(a.value);
      const bNum = parseFloat(b.value);

      const aValid = !isNaN(aNum);
      const bValid = !isNaN(bNum);

      if (!aValid && !bValid) {
        cmp = 0;
      } else if (!aValid) {
        cmp = 1;
      } else if (!bValid) {
        cmp = -1;
      } else {
        cmp = aNum - bNum;
      }
    } else {
      const aStr = (a.value || "").toLocaleLowerCase();
      const bStr = (b.value || "").toLocaleLowerCase();
      cmp = aStr.localeCompare(bStr);
    }

    if (cmp === 0) {
      return a.originalIndex - b.originalIndex;
    }

    return ascending ? cmp : -cmp;
  });

  rowData.forEach(item => tbody.appendChild(item.row));
}

function detectNumericColumn(rowData, table, columnIndex) {
  // Headers with scope='col' are the sortable ones (Client, Name, Email, Cc)
  const sortableHeaders = table.querySelectorAll("thead th[scope='col']");
  const header = sortableHeaders[columnIndex];

  if (header && (header.classList.contains("is-num") || header.classList.contains("num"))) {
    return true;
  }

  // Account for optional delete/actions column at the start
  const hasDeleteColumn = !!table.querySelector("th.rm-actions-header");
  const cellIndex = hasDeleteColumn ? columnIndex + 1 : columnIndex;

  const bodyCells = table.querySelectorAll(
    "tbody tr td:nth-child(" + (cellIndex + 1) + ")"
  );

  for (const cell of bodyCells) {
    if (cell.classList.contains("is-num") || cell.classList.contains("num")) {
      return true;
    }
  }

  const numericThreshold = 0.7;
  let numericCount = 0;
  let nonEmptyCount = 0;

  rowData.forEach(item => {
    if (item.value !== "") {
      nonEmptyCount++;
      const num = parseFloat(item.value);
      if (!isNaN(num)) numericCount++;
    }
  });

  if (nonEmptyCount === 0) return false;

  return numericCount / nonEmptyCount >= numericThreshold;
}

console.log("[RM Content] Script initialization complete");

} // end of else block (guard against multiple injections)
