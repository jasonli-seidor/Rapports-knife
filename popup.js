/**
 * Configuration object for the Jira Worklog Reporter extension.
 */
const CONFIG = {
  URLS: {
    RAPPORTS: "https://intranetnew.seidor.com/rapports/imputation-hours",
    RAPPORTS_API_BASE: "https://apis-intranet.seidor.com",
    JIRA_API_BASE: "https://seidorcc.atlassian.net/rest/api/3",
    SEIDOR_INTRANET: "https://intranetnew.seidor.com/*",
  },

  API_ENDPOINTS: {
    USER_PROFILE: "/authorizationv2/user-profile",
    PROJECTS: "/authorizationv2/paginated-projects",
    SUB_PROJECTS: "/collections/paginated-subprojects",
    IMPUTATIONS: "/rapports/imputations",
    JIRA_MYSELF: "/myself",
    JIRA_SEARCH: "/search",
    JIRA_ISSUE_WORKLOG: (issueId) => `/issue/${issueId}/worklog`,
  },

  RAPPORTS_PAYLOAD: {
    CATEGORY: "PR",
    SITUATION_ID: "6", // Situation: office
    DEFAULT_TASK_ID: "",
    INTERNAL_REF: "",
  },

  JIRA: {
    PEP_CUSTOM_FIELD: "customfield_10120",
    PEP_MAPPING: {
      // Key: Jira issue key prefix or full key
      // Value: Rule for mapping
      "LEC-": {
        // If the PEP value includes "14-SEIDOR-AM" or is empty, set the PEP to "14-SEIDOR-AM&LEC"
        condition: (pep) =>
          pep.value.toUpperCase().includes("14-SEIDOR-AM") ||
          pep.value.length === 0,
        result: { pep: { value: "14-SEIDOR-AM&LEC" } },
      },
      "SA-17": {
        result: { pep: { value: "14-ZPR-VAC25" } },
      },
      "SA-18": {
        result: {
          pep: { value: "14-SEIDOR-AM&GENERAL" },
          comment: "Daily Standup",
        },
      },
      "SA-19": {
        // If the comment includes "team building" or "teambuilding", set the PEP to "14-ZPR-TA&TEAMBUILDING"
        condition: (_pep, comment) =>
          comment.toLowerCase().includes("team building") ||
          comment.toLowerCase().includes("teambuilding"),
        result: { pep: { value: "14-ZPR-TA&TEAMBUILDING" } },
        fallback: { pep: { value: "14-ZPR-TA&OTHERS" } },
      },
    },
  },
};

// --- DOM Elements ---
const getElements = () => ({
  startDateInput: document.getElementById("startDate"),
  endDateInput: document.getElementById("endDate"),
  getWorklogsButton: document.getElementById("getWorklogs"),
  getUserProfileButton: document.getElementById("getUserProfile"),
  getProjectsButton: document.getElementById("getProjects"),
  syncWorklogsButton: document.getElementById("syncWorklogs"),
  statusDiv: document.getElementById("status"),
  devToolsToggle: document.getElementById("devToolsToggle"),
  devTools: document.getElementById("devTools"),
  syncHintDiv: document.getElementById("syncHint"),
  modal: document.getElementById("subProjectModal"),
  modalTitle: document.getElementById("modalTitle"),
  choicesDiv: document.getElementById("subProjectChoices"),
  confirmBtn: document.getElementById("confirmSubProject"),
  cancelBtn: document.getElementById("cancelSubProject"),
});

// --- API Layer ---

async function fetchFromApi(url, options = {}, token) {
  const headers = {
    ...options.headers,
    "Content-Type": "application/json",
  };
  if (token) {
    headers["Authorization"] = `Bearer ${token}`;
  }
  const response = await fetch(url, { ...options, headers });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(
      `API request to ${url} failed with status ${response.status}: ${errorBody}`
    );
  }
  return response.json();
}

async function getBearerToken() {
  const queryOptions = { url: CONFIG.URLS.SEIDOR_INTRANET };
  let [activeTab] = await chrome.tabs.query({ ...queryOptions, active: true });
  const intranetTab =
    activeTab || (await chrome.tabs.query(queryOptions)).at(0);

  if (!intranetTab) {
    throw new Error(
      "Please open and log in to the Seidor Rapports in another tab."
    );
  }

  const results = await chrome.scripting.executeScript({
    target: { tabId: intranetTab.id },
    function: () => sessionStorage.getItem("appState"),
  });

  if (!results?.[0]?.result) {
    throw new Error("Could not retrieve 'appState' from session storage.");
  }
  const appState = JSON.parse(results[0].result);
  const token = appState?.tokenData?.accessToken;

  if (typeof token !== "string" || !token.startsWith("ey")) {
    throw new Error("Could not find a valid token within appState.");
  }
  return token;
}

const rapportApi = {
  getUserData: (token) =>
    fetchFromApi(
      CONFIG.URLS.RAPPORTS_API_BASE + CONFIG.API_ENDPOINTS.USER_PROFILE,
      {},
      token
    ),

  getProjectsData: (token) =>
    fetchFromApi(
      CONFIG.URLS.RAPPORTS_API_BASE + CONFIG.API_ENDPOINTS.PROJECTS,
      {
        method: "POST",
        body: JSON.stringify({
          multiSortedColumns: [{ active: "label", direction: "asc" }],
          filterMap: { moduleId: "2" },
          pagination: { pageNumber: 1, pageSize: 100000 },
        }),
      },
      token
    ),

  getSubProjectsData: (projectId, token) =>
    fetchFromApi(
      CONFIG.URLS.RAPPORTS_API_BASE + CONFIG.API_ENDPOINTS.SUB_PROJECTS,
      {
        method: "POST",
        body: JSON.stringify({
          filterMap: { projectId, isActive: "true" },
          pagination: { pageNumber: 1, pageSize: 100000 },
        }),
      },
      token
    ),

  postImputation: (payload, token) =>
    fetchFromApi(
      CONFIG.URLS.RAPPORTS_API_BASE + CONFIG.API_ENDPOINTS.IMPUTATIONS,
      { method: "POST", body: JSON.stringify(payload) },
      token
    ),
};

const jiraApi = {
  fetch: (endpoint, options = {}) => {
    return fetch(CONFIG.URLS.JIRA_API_BASE + endpoint, options);
  },

  getMyself: async () => {
    const response = await jiraApi.fetch(CONFIG.API_ENDPOINTS.JIRA_MYSELF);
    if (!response.ok) throw new Error("Failed to fetch Jira user.");
    return response.json();
  },

  searchIssuesWithWorklogs: async (jql) => {
    const searchUrl = `${
      CONFIG.API_ENDPOINTS.JIRA_SEARCH
    }?jql=${encodeURIComponent(jql)}&fields=${CONFIG.JIRA.PEP_CUSTOM_FIELD}`;
    const response = await jiraApi.fetch(searchUrl);
    if (!response.ok) throw new Error("Failed to search Jira issues.");
    const searchData = await response.json();
    return searchData.issues || [];
  },

  getIssueWorklogs: async (issueId) => {
    const response = await jiraApi.fetch(
      CONFIG.API_ENDPOINTS.JIRA_ISSUE_WORKLOG(issueId)
    );
    if (!response.ok) return []; // Gracefully fail for single issue
    const result = await response.json();
    return result.worklogs || [];
  },
};

// --- Logic ---

function customizeWorklogDetails(issue, originalComment) {
  const { PEP_MAPPING, PEP_CUSTOM_FIELD } = CONFIG.JIRA;
  const issuePep = issue.fields[PEP_CUSTOM_FIELD];
  let finalPep = issuePep;
  let finalComment = originalComment;

  for (const key in PEP_MAPPING) {
    if (issue.key.startsWith(key)) {
      const rule = PEP_MAPPING[key];
      const conditionMet = rule.condition
        ? rule.condition(issuePep, originalComment)
        : true;

      if (conditionMet) {
        if (rule.result.pep) finalPep = rule.result.pep;
        if (rule.result.comment && originalComment === "No comment") {
          finalComment = rule.result.comment;
        }
      } else if (rule.fallback) {
        if (rule.fallback.pep) finalPep = rule.fallback.pep;
      }
      // example comment: [FA-20] something something
      finalComment = `[${issue.key}] \n${finalComment}`;
      break; // Assume first matching rule is sufficient
    }
  }

  return { pep: finalPep, comment: finalComment };
}

async function getJiraWorklogs(startDate, endDate) {
  const { accountId } = await jiraApi.getMyself();
  const jql = `worklogDate >= "${startDate}" AND worklogDate <= "${endDate}" AND worklogAuthor = currentUser()`;

  const issues = await jiraApi.searchIssuesWithWorklogs(jql);
  if (issues.length === 0) return [];

  const worklogPromises = issues.map(async (issue) => {
    const worklogsForIssue = await jiraApi.getIssueWorklogs(issue.id);
    return worklogsForIssue
      .filter((wl) => {
        const worklogDate = new Date(wl.started).toISOString().split("T")[0];
        return (
          wl.author.accountId === accountId &&
          worklogDate >= startDate &&
          worklogDate <= endDate
        );
      })
      .map((worklog) => {
        const commentText =
          worklog.comment?.content?.[0]?.content?.[0]?.text || "No comment";
        const { pep, comment } = customizeWorklogDetails(issue, commentText);
        return {
          timeSpentSeconds: worklog.timeSpentSeconds,
          comment,
          started: worklog.started,
          pep,
        };
      });
  });

  const nestedWorklogs = await Promise.all(worklogPromises);
  return nestedWorklogs.flat();
}

const formatters = {
  date: (date) =>
    `${String(date.getDate()).padStart(2, "0")}/${String(
      date.getMonth() + 1
    ).padStart(2, "0")}/${date.getFullYear()}`,
  hours: (seconds) => {
    const totalMinutes = Math.floor(seconds / 60);
    const hours = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
    const minutes = String(totalMinutes % 60).padStart(2, "0");
    return `${hours}:${minutes}`;
  },
};

// --- Version Check ---

function isNewerVersion(remote, local) {
  const remoteParts = remote.split(".").map(Number);
  const localParts = local.split(".").map(Number);

  for (let i = 0; i < Math.max(remoteParts.length, localParts.length); i++) {
    const remotePart = remoteParts[i] || 0;
    const localPart = localParts[i] || 0;
    if (remotePart > localPart) return true;
    if (remotePart < localPart) return false;
  }
  return false;
}

function displayUpdateHint(newVersion) {
  let updateHintDiv = document.getElementById("updateHint");
  if (!updateHintDiv) {
    updateHintDiv = document.createElement("div");
    updateHintDiv.id = "updateHint";
    updateHintDiv.style.padding = "10px";
    updateHintDiv.style.backgroundColor = "#fff3cd";
    updateHintDiv.style.color = "#664d03";
    updateHintDiv.style.border = "1px solid #ffc107";
    updateHintDiv.style.borderRadius = "5px";
    updateHintDiv.style.marginBottom = "10px";
    updateHintDiv.style.textAlign = "center";
    const firstChild = document.body.firstChild;
    document.body.insertBefore(updateHintDiv, firstChild);
  }

  updateHintDiv.innerHTML = `
        New version ${newVersion} is available,<br>please pull the latest changes.
        <button id="dismissUpdate" style="margin-left: 10px; border: none; background: transparent; color: #6c757d; cursor: pointer; text-decoration: underline; font-size: 12px; padding: 0;">Don't show again</button>
    `;

  document.getElementById("dismissUpdate").addEventListener("click", () => {
    chrome.storage.local.set({ dismissedVersion: newVersion });
    updateHintDiv.style.display = "none";
  });
}

async function checkVersion() {
  try {
    const remoteManifestUrl =
      "https://raw.githubusercontent.com/jaysonliang-seidor/Rapports-knife/refs/heads/main/manifest.json";
    const response = await fetch(remoteManifestUrl);
    if (!response.ok) {
      console.warn("Could not fetch remote manifest for version check.");
      return;
    }
    const remoteManifest = await response.json();
    const remoteVersion = remoteManifest.version;

    const localManifest = chrome.runtime.getManifest();
    const localVersion = localManifest.version;

    if (isNewerVersion(remoteVersion, localVersion)) {
      chrome.storage.local.get("dismissedVersion", ({ dismissedVersion }) => {
        if (dismissedVersion !== remoteVersion) {
          displayUpdateHint(remoteVersion);
        }
      });
    }
  } catch (error) {
    console.error("Version check failed:", error);
  }
}

// --- UI Functions ---

function updateStatus(message, isHtml = false) {
  const { statusDiv } = getElements();
  if (isHtml) {
    statusDiv.innerHTML = message;
  } else {
    statusDiv.textContent = message;
  }
}

async function updateButtonStates(elements) {
  const {
    startDateInput,
    endDateInput,
    syncWorklogsButton,
    getWorklogsButton,
    syncHintDiv,
    getUserProfileButton,
    getProjectsButton,
  } = elements;
  const datesSelected = startDateInput.value && endDateInput.value;
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const onCorrectPage = tab?.url?.startsWith(CONFIG.URLS.RAPPORTS);

  const startDate = new Date(startDateInput.value + "T00:00:00");
  const today = new Date();
  const isPreviousMonth =
    startDate.getFullYear() < today.getFullYear() ||
    (startDate.getFullYear() === today.getFullYear() &&
      startDate.getMonth() < today.getMonth());

  // This button only depends on having dates selected
  getWorklogsButton.disabled = !datesSelected;

  // These buttons depend on being on the Rapports page
  const hint = "Disabled outside of Rapports";
  getUserProfileButton.disabled = !onCorrectPage;
  getUserProfileButton.title = !onCorrectPage ? hint : "";

  getProjectsButton.disabled = !onCorrectPage;
  getProjectsButton.title = !onCorrectPage ? hint : "";

  // The sync button has multiple conditions
  syncWorklogsButton.disabled =
    !datesSelected || !onCorrectPage || isPreviousMonth;
  syncWorklogsButton.title = !onCorrectPage ? hint : "";

  if (!onCorrectPage) {
    syncHintDiv.innerHTML = `Please go to <a href="${CONFIG.URLS.RAPPORTS}">Rapports</a> to activate the sync feature.`;
    syncHintDiv.querySelector("a")?.addEventListener("click", (e) => {
      e.preventDefault();
      chrome.tabs.create({ url: e.target.href });
      window.close();
    });
  } else {
    syncHintDiv.innerHTML = "";
  }
}

function promptForSubProject(matches, keyword) {
  const { modal, choicesDiv, confirmBtn, cancelBtn, modalTitle } =
    getElements();

  modalTitle.textContent = `Select Sub-Project for "${keyword}"`;
  choicesDiv.innerHTML = ""; // Clear previous choices

  return new Promise((resolve, reject) => {
    matches.forEach((match, index) => {
      const radioId = `subproject-choice-${index}`;
      const wrapper = document.createElement("div");
      const input = document.createElement("input");
      input.type = "radio";
      input.name = "subProjectChoice";
      input.id = radioId;
      input.value = match.value;
      if (index === 0) input.checked = true;

      const label = document.createElement("label");
      label.htmlFor = radioId;
      label.textContent = match.label;

      wrapper.appendChild(input);
      wrapper.appendChild(label);
      choicesDiv.appendChild(wrapper);
    });

    const cleanup = (listener) => {
      modal.style.display = "none";
      confirmBtn.removeEventListener("click", onConfirm);
      cancelBtn.removeEventListener("click", onCancel);
    };

    const onConfirm = () => {
      cleanup();
      const selected = choicesDiv.querySelector("input:checked");
      resolve(selected.value);
    };

    const onCancel = () => {
      cleanup();
      reject(new Error("User cancelled selection."));
    };

    confirmBtn.addEventListener("click", onConfirm);
    cancelBtn.addEventListener("click", onCancel);
    modal.style.display = "flex";
  });
}

function displaySyncSummary(successCount, failedLogs) {
  const failureCount = failedLogs.length;
  let summaryMessage = `Sync complete. Success: ${successCount}, Failed: ${failureCount}.`;

  if (failureCount > 0) {
    let failureHtml = `
      <div style="margin-top: 10px; font-style: normal; text-align: left;">
        <strong>Failed Syncs:</strong>
        <ul style="padding-left: 20px; margin-top: 5px; max-height: 100px; overflow-y: auto;">`;
    for (const log of failedLogs) {
      const pep = log.pep.replace(/</g, "&lt;").replace(/>/g, "&gt;");
      const reason = log.reason.replace(/</g, "&lt;").replace(/>/g, "&gt;");
      failureHtml += `<li style="margin-bottom: 5px;"><strong>${log.date}</strong> (${pep}):<br>${reason}</li>`;
    }
    failureHtml += "</ul></div>";
    updateStatus(summaryMessage + failureHtml, true);
    console.warn("--- Failed Syncs ---", failedLogs);
  } else {
    updateStatus(summaryMessage + " All worklogs synced. Reloading page...");
    setTimeout(async () => {
      const [tab] = await chrome.tabs.query({
        active: true,
        url: `${CONFIG.URLS.RAPPORTS}*`,
      });
      if (tab) chrome.tabs.reload(tab.id);
    }, 1000);
  }
}

// --- Event Handlers ---

async function handleDevAction(action, successMessage) {
  updateStatus(`Fetching ${action.name}...`);
  try {
    const token = await getBearerToken();
    const result = await action(token);
    console.log(`--- ${action.name} ---`, result);
    updateStatus(successMessage);
  } catch (error) {
    console.error("Error:", error);
    updateStatus(`Error: ${error.message}`);
  }
}

async function handleFetchWorklogs() {
  updateStatus("Fetching worklogs...");
  const { startDateInput, endDateInput } = getElements();
  try {
    const worklogs = await getJiraWorklogs(
      startDateInput.value,
      endDateInput.value
    );
    console.log("--- Filtered Worklogs ---", worklogs);
    updateStatus(`Found and logged ${worklogs.length} worklog(s) to console.`);
  } catch (error) {
    console.error("Error:", error);
    updateStatus(`Error: ${error.message}`);
  }
}

async function handleSyncWorklogs() {
  const { startDateInput, endDateInput } = getElements();

  updateStatus("Starting sync...");
  try {
    const token = await getBearerToken();
    updateStatus("Fetching all required data...");

    const [user, projectsResponse, worklogs] = await Promise.all([
      rapportApi.getUserData(token),
      rapportApi.getProjectsData(token),
      getJiraWorklogs(startDateInput.value, endDateInput.value),
    ]);

    if (worklogs.length === 0) {
      updateStatus("No worklogs found in Jira for the selected period.");
      return;
    }

    const projects = projectsResponse.data;
    const projectMap = projects.reduce((acc, proj) => {
      acc[proj.label] = proj.value;
      return acc;
    }, {});

    let successCount = 0;
    const failedLogs = [];

    updateStatus(
      `Data fetched. Starting imputation for ${worklogs.length} logs...`
    );

    for (const [index, worklog] of worklogs.entries()) {
      const worklogDateStr = new Date(worklog.started)
        .toISOString()
        .split("T")[0];
      const pepValue = worklog.pep?.value;

      const fail = (reason) => {
        failedLogs.push({
          date: worklogDateStr,
          pep: pepValue || "N/A",
          reason,
        });
      };

      if (!pepValue) {
        fail("Missing PEP value in Jira.");
        continue;
      }

      let projectLabel = pepValue;
      let subProjectKeyword = null;
      if (pepValue.includes("&")) {
        [projectLabel, subProjectKeyword] = pepValue.split("&");
      }

      const projectId = projectMap[projectLabel];
      if (!projectId) {
        fail(`Project "${projectLabel}" not found in Rapports.`);
        continue;
      }

      let subProjectId = CONFIG.RAPPORTS_PAYLOAD.DEFAULT_TASK_ID;
      if (subProjectKeyword) {
        try {
          const subProjects = (
            await rapportApi.getSubProjectsData(projectId, token)
          ).data;
          const matching = subProjects.filter((sp) =>
            sp.label.toUpperCase().includes(subProjectKeyword.toUpperCase())
          );

          if (matching.length === 1) {
            subProjectId = matching[0].value;
          } else if (matching.length > 1) {
            updateStatus(
              `Waiting for sub-project selection for PEP: ${pepValue}`
            );
            subProjectId = await promptForSubProject(
              matching,
              subProjectKeyword
            );
          } else {
            fail(`Sub-project with keyword "${subProjectKeyword}" not found.`);
            continue;
          }
        } catch (error) {
          fail(
            error.message.includes("cancelled")
              ? "Sub-project selection cancelled."
              : "API error fetching sub-projects."
          );
          continue;
        }
      }

      const date = new Date(worklog.started);
      const payload = {
        id: "",
        fromDate: formatters.date(date),
        toDate: formatters.date(date),
        userId: user.id,
        projectId,
        subProjectId,
        description: worklog.comment,
        hours: formatters.hours(worklog.timeSpentSeconds),
        category: CONFIG.RAPPORTS_PAYLOAD.CATEGORY,
        situationId: CONFIG.RAPPORTS_PAYLOAD.SITUATION_ID,
        taskId: CONFIG.RAPPORTS_PAYLOAD.DEFAULT_TASK_ID,
        internalRef: CONFIG.RAPPORTS_PAYLOAD.INTERNAL_REF,
      };

      try {
        updateStatus(
          `Syncing ${index + 1}/${worklogs.length}... (PEP: ${pepValue})`
        );
        await rapportApi.postImputation(payload, token);
        successCount++;
      } catch (error) {
        fail(`Imputation API call failed: ${error.message}`);
      }
    }
    displaySyncSummary(successCount, failedLogs);
  } catch (error) {
    console.error("Sync failed:", error);
    updateStatus(`Error: ${error.message}`);
  }
}

// --- Initializer ---

document.addEventListener("DOMContentLoaded", () => {
  const elements = getElements();

  const {
    startDateInput,
    endDateInput,
    devToolsToggle,
    getWorklogsButton,
    getUserProfileButton,
    getProjectsButton,
    syncWorklogsButton,
    devTools,
  } = elements;

  // Set default dates
  const today = new Date();
  const todayString = today.toISOString().split("T")[0];
  startDateInput.value = todayString;
  endDateInput.value = todayString;

  // Restore Dev Tools state
  chrome.storage.local.get("devToolsVisible", ({ devToolsVisible }) => {
    if (devToolsVisible) {
      devTools.style.display = "block";
      devToolsToggle.checked = true;
    }
  });

  // Add Event Listeners
  startDateInput.addEventListener("change", () => updateButtonStates(elements));
  endDateInput.addEventListener("change", () => updateButtonStates(elements));

  devToolsToggle.addEventListener("change", () => {
    const isVisible = devToolsToggle.checked;
    devTools.style.display = isVisible ? "block" : "none";
    chrome.storage.local.set({ devToolsVisible: isVisible });
  });

  getWorklogsButton.addEventListener("click", handleFetchWorklogs);
  getUserProfileButton.addEventListener("click", () =>
    handleDevAction(rapportApi.getUserData, "User profile logged to console.")
  );
  getProjectsButton.addEventListener("click", () =>
    handleDevAction(rapportApi.getProjectsData, "Projects logged to console.")
  );
  syncWorklogsButton.addEventListener("click", handleSyncWorklogs);

  checkVersion();
  updateButtonStates(elements);
});
