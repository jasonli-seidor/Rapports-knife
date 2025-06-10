document.addEventListener("DOMContentLoaded", async () => {
  const startDateInput = document.getElementById("startDate");
  const endDateInput = document.getElementById("endDate");
  const getWorklogsButton = document.getElementById("getWorklogs");
  const getUserProfileButton = document.getElementById("getUserProfile");
  const getProjectsButton = document.getElementById("getProjects");
  const syncWorklogsButton = document.getElementById("syncWorklogs");
  const statusDiv = document.getElementById("status");
  const devToolsToggle = document.getElementById("devToolsToggle");
  const devTools = document.getElementById("devTools");
  const syncHintDiv = document.getElementById("syncHint");

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const rapportsUrl =
    "https://intranetnew.seidor.com/rapports/imputation-hours";
  const onCorrectPage = tab?.url?.startsWith(rapportsUrl);

  function updateButtonStates() {
    const datesSelected = startDateInput.value && endDateInput.value;

    getWorklogsButton.disabled = !datesSelected;
    syncWorklogsButton.disabled = !datesSelected || !onCorrectPage;

    if (!onCorrectPage) {
      syncHintDiv.innerHTML = `Please go to <a href="${rapportsUrl}">Rapports</a> to activate the sync feature.`;
      const link = syncHintDiv.querySelector("a");
      if (link) {
        link.addEventListener("click", (e) => {
          e.preventDefault();
          chrome.tabs.create({ url: e.target.href });
          window.close();
        });
      }
    } else {
      syncHintDiv.innerHTML = "";
    }
  }

  startDateInput.addEventListener("change", updateButtonStates);
  endDateInput.addEventListener("change", updateButtonStates);

  // --- Dev Tools Toggle ---
  chrome.storage.local.get("devToolsVisible", (data) => {
    if (data.devToolsVisible) {
      devTools.style.display = "block";
      devToolsToggle.checked = true;
    }
  });

  devToolsToggle.addEventListener("change", () => {
    const isVisible = devToolsToggle.checked;
    devTools.style.display = isVisible ? "block" : "none";
    chrome.storage.local.set({ devToolsVisible: isVisible });
  });

  // --- Main Event Listeners ---
  getWorklogsButton.addEventListener("click", handleFetchWorklogs);
  getUserProfileButton.addEventListener("click", handleFetchUserProfile);
  getProjectsButton.addEventListener("click", handleFetchProjects);
  syncWorklogsButton.addEventListener("click", handleSyncWorklogs);

  updateButtonStates();

  // --- Reusable Core Data Fetching Functions ---

  async function getBearerToken() {
    // Prioritize finding an active tab matching the URL.
    let tabs = await chrome.tabs.query({
      active: true,
      url: "https://intranetnew.seidor.com/*",
    });

    // If no active tab is found, search for any tab with the URL.
    if (tabs.length === 0) {
      tabs = await chrome.tabs.query({
        url: "https://intranetnew.seidor.com/*",
      });
    }

    if (tabs.length === 0)
      throw new Error(
        "Please open and log in to the Seidor Rapports in another tab."
      );

    const intranetTab = tabs[0];
    const results = await chrome.scripting.executeScript({
      target: { tabId: intranetTab.id },
      function: () => sessionStorage.getItem("appState"),
    });

    if (!results || results.length === 0 || !results[0].result)
      throw new Error("Could not retrieve 'appState' from session storage.");

    const appState = JSON.parse(results[0].result);
    if (appState?.tokenData?.accessToken) {
      const token = appState.tokenData.accessToken;
      if (
        typeof token === "string" &&
        token.startsWith("ey") &&
        token.length > 100
      )
        return token;
    }
    throw new Error("Could not find a valid token within appState.");
  }

  async function getUserData(token) {
    const response = await fetch(
      "https://apis-intranet.seidor.com/authorizationv2/user-profile",
      {
        headers: { authorization: `Bearer ${token}` },
      }
    );
    if (!response.ok)
      throw new Error(`Fetching user profile failed: ${response.statusText}`);
    return await response.json();
  }

  async function getProjectsData(token) {
    const response = await fetch(
      "https://apis-intranet.seidor.com/authorizationv2/paginated-projects",
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${token}`,
          "content-type": "application/json",
        },
        body: JSON.stringify({
          multiSortedColumns: [{ active: "label", direction: "asc" }],
          filterMap: { moduleId: "2" },
          pagination: { pageNumber: 1, pageSize: 100000 },
        }),
      }
    );
    if (!response.ok)
      throw new Error(`Fetching projects failed: ${response.statusText}`);
    const projects = await response.json();
    return projects.data;
  }

  async function getSubProjectsData(projectId, token) {
    const response = await fetch(
      "https://apis-intranet.seidor.com/collections/paginated-subprojects",
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${token}`,
          "content-type": "application/json",
        },
        body: JSON.stringify({
          filterMap: {
            projectId: projectId,
            isActive: "true",
          },
          pagination: {
            pageNumber: 1,
            pageSize: 100000,
          },
        }),
      }
    );
    if (!response.ok)
      throw new Error(
        `Fetching sub-projects for project ID ${projectId} failed: ${response.statusText}`
      );
    const subProjects = await response.json();
    return subProjects.data;
  }

  async function getJiraWorklogs(startDate, endDate) {
    const JIRA_BASE_URL = "https://seidorcc.atlassian.net";
    const userResponse = await fetch(`${JIRA_BASE_URL}/rest/api/3/myself`);
    if (!userResponse.ok)
      throw new Error(`Failed to fetch Jira user: ${userResponse.statusText}`);
    const userData = await userResponse.json();
    const accountId = userData.accountId;

    const jql = `worklogDate >= "${startDate}" AND worklogDate <= "${endDate}" AND worklogAuthor = currentUser()`;
    const searchUrl = `${JIRA_BASE_URL}/rest/api/3/search?jql=${encodeURIComponent(
      jql
    )}&fields=customfield_10120`;
    const searchResponse = await fetch(searchUrl);
    if (!searchResponse.ok)
      throw new Error(
        `Failed to search Jira issues: ${searchResponse.statusText}`
      );
    const searchData = await searchResponse.json();
    const issues = searchData.issues;

    if (issues.length === 0) return [];

    const issueInfoMap = issues.reduce((acc, issue) => {
      acc[issue.id] = { pep: issue.fields.customfield_10120, key: issue.key };
      return acc;
    }, {});

    const worklogPromises = issues.map((issue) =>
      fetch(`${JIRA_BASE_URL}/rest/api/3/issue/${issue.id}/worklog`).then(
        (res) => res.json()
      )
    );
    const worklogResults = await Promise.all(worklogPromises);
    const allWorklogs = worklogResults.flatMap((result) => result.worklogs);

    return allWorklogs
      .filter((worklog) => {
        const worklogStartedDate = new Date(worklog.started)
          .toISOString()
          .split("T")[0];
        return (
          worklog.author.accountId === accountId &&
          worklogStartedDate >= startDate &&
          worklogStartedDate <= endDate
        );
      })
      .map((worklog) => {
        let commentText =
          worklog.comment?.content?.[0]?.content?.[0]?.text || "No comment";
        let issueInfo = issueInfoMap[worklog.issueId] || {};

        // Custom PEP logic as if/else statements
        if (
          issueInfo.key.startsWith("LEC-") &&
          (issueInfo.pep.value.includes("14-SEIDOR-AM".toUpperCase()) ||
            issueInfo.pep.value.length === 0)
        ) {
          issueInfo.pep = { value: "14-SEIDOR-AM&LEC" };
        }
        if (issueInfo.key === "SA-17") {
          issueInfo.pep = { value: "14-ZPR-VAC25" };
        }
        if (issueInfo.key === "SA-18") {
          issueInfo.pep = { value: "14-SEIDOR-AM&GENERAL" };
          if (commentText === "No comment") {
            commentText = "Daily Standup";
          }
        }
        if (issueInfo.key === "SA-19") {
          if (
            commentText.toLowerCase().includes("team building") ||
            commentText.toLowerCase().includes("teambuilding")
          ) {
            issueInfo.pep = { value: "14-ZPR-TA&TEAMBUILDING" };
          } else {
            issueInfo.pep = { value: "14-ZPR-TA&OTHERS" };
          }
        }

        return {
          timeSpentSeconds: worklog.timeSpentSeconds,
          comment: commentText,
          started: worklog.started,
          PEP: issueInfo.pep,
        };
      });
  }

  async function postImputation(payload, token) {
    const response = await fetch(
      "https://apis-intranet.seidor.com/rapports/imputations",
      {
        method: "POST",
        headers: {
          accept: "application/json, text/plain, */*",
          authorization: `Bearer ${token}`,
          "content-type": "application/json",
        },
        body: JSON.stringify(payload),
      }
    );
    if (!response.ok) {
      const errorBody = await response.text();
      throw new Error(`Imputation failed for project ${payload.projectId} 
                on ${payload.fromDate}: ${response.statusText} - ${errorBody}`);
    }
    return await response.json();
  }

  async function handleFetchUserProfile() {
    statusDiv.textContent = "Fetching user profile...";
    try {
      const token = await getBearerToken();
      const userProfile = await getUserData(token);
      console.log("--- User Profile ---", userProfile);
      statusDiv.textContent = "User profile logged to console.";
    } catch (error) {
      console.error("Error:", error);
      statusDiv.textContent = `Error: ${error.message}`;
    }
  }

  async function handleFetchProjects() {
    statusDiv.textContent = "Fetching projects...";
    try {
      const token = await getBearerToken();
      const projects = await getProjectsData(token);
      console.log("--- Projects ---", projects);
      statusDiv.textContent = "Projects logged to console.";
    } catch (error) {
      console.error("Error:", error);
      statusDiv.textContent = `Error: ${error.message}`;
    }
  }

  async function handleFetchWorklogs() {
    statusDiv.textContent = "Fetching worklogs...";
    try {
      const worklogs = await getJiraWorklogs(
        startDateInput.value,
        endDateInput.value
      );
      console.log("--- Filtered Worklogs ---", worklogs);
      statusDiv.textContent = `Found and logged ${worklogs.length} worklog(s) to the console.`;
    } catch (error) {
      console.error("Error:", error);
      statusDiv.textContent = `Error: ${error.message}`;
    }
  }

  function promptForSubProject(matches, keyword) {
    const modal = document.getElementById("subProjectModal");
    const choicesDiv = document.getElementById("subProjectChoices");
    const confirmBtn = document.getElementById("confirmSubProject");
    const cancelBtn = document.getElementById("cancelSubProject");
    const modalTitle = document.getElementById("modalTitle");

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

      const onConfirm = () => {
        cleanup();
        const selected = choicesDiv.querySelector("input:checked");
        resolve(selected.value);
      };

      const onCancel = () => {
        cleanup();
        reject(new Error("User cancelled selection."));
      };

      const cleanup = () => {
        modal.style.display = "none";
        confirmBtn.removeEventListener("click", onConfirm);
        cancelBtn.removeEventListener("click", onCancel);
      };

      confirmBtn.addEventListener("click", onConfirm);
      cancelBtn.addEventListener("click", onCancel);
      modal.style.display = "flex";
    });
  }

  async function handleSyncWorklogs() {
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();

    // The time zone offset needs to be considered for an accurate date comparison.
    const startDate = new Date(startDateInput.value + "T00:00:00");
    const startMonth = startDate.getMonth();
    const startYear = startDate.getFullYear();

    if (
      startYear < currentYear ||
      (startYear === currentYear && startMonth < currentMonth)
    ) {
      statusDiv.innerHTML =
        '<strong style="color: red;">Syncing for previous months is not allowed.</strong>';
      return;
    }

    statusDiv.textContent = "Starting sync...";
    try {
      const token = await getBearerToken();
      statusDiv.textContent = "Fetching all required data...";

      const [user, projects, worklogs] = await Promise.all([
        getUserData(token),
        getProjectsData(token),
        getJiraWorklogs(startDateInput.value, endDateInput.value),
      ]);

      if (worklogs.length === 0) {
        statusDiv.textContent = "No worklogs found in Jira to sync.";
        return;
      }

      const projectMap = projects.reduce((acc, proj) => {
        acc[proj.label] = proj.value;
        return acc;
      }, {});

      let successCount = 0;
      let failureCount = 0;
      const failedLogs = [];
      const total = worklogs.length;

      statusDiv.textContent = `Data fetched. Starting imputation for ${total} worklogs...`;

      for (const [index, worklog] of worklogs.entries()) {
        const worklogDate = new Date(worklog.started)
          .toISOString()
          .split("T")[0];
        const pepValue = worklog.PEP?.value;
        if (!pepValue) {
          console.warn("Skipping worklog due to missing PEP value:", worklog);
          failureCount++;
          failedLogs.push({
            date: worklogDate,
            pep: "N/A",
            reason: "Missing PEP value in Jira.",
          });
          continue;
        }

        let projectLabel = pepValue;
        let subProjectKeyword = null;
        let subProjectId = "";

        if (pepValue.includes("&")) {
          const parts = pepValue.split("&");
          projectLabel = parts[0];
          subProjectKeyword = parts[1];
        }

        const projectId = projectMap[projectLabel];
        if (!projectId) {
          console.warn(
            `Skipping worklog for project label "${projectLabel}" - no matching project ID found in Rapports.`,
            worklog
          );
          failureCount++;
          failedLogs.push({
            date: worklogDate,
            pep: pepValue,
            reason: `Project "${projectLabel}" not found in Rapports.`,
          });
          continue;
        }

        if (subProjectKeyword) {
          try {
            const subProjects = await getSubProjectsData(projectId, token);
            const matchingSubProjects = subProjects.filter((sp) =>
              sp.label.toUpperCase().includes(subProjectKeyword.toUpperCase())
            );

            if (matchingSubProjects.length === 1) {
              subProjectId = matchingSubProjects[0].value;
            } else if (matchingSubProjects.length > 1) {
              try {
                statusDiv.textContent = `Waiting for selection for PEP: ${pepValue}`;
                subProjectId = await promptForSubProject(
                  matchingSubProjects,
                  subProjectKeyword
                );
                statusDiv.textContent = `Syncing ${index + 1}/${total}...`;
              } catch (selectionError) {
                console.warn(
                  `Skipping due to user cancellation for PEP: ${pepValue}`
                );
                failureCount++;
                failedLogs.push({
                  date: worklogDate,
                  pep: pepValue,
                  reason: "Sub-project selection cancelled.",
                });
                continue;
              }
            } else {
              console.warn(
                `Skipping: Found project "${projectLabel}" but couldn't find sub-project with keyword "${subProjectKeyword}".`,
                worklog
              );
              failureCount++;
              failedLogs.push({
                date: worklogDate,
                pep: pepValue,
                reason: `Sub-project keyword "${subProjectKeyword}" not found.`,
              });
              continue;
            }
          } catch (subProjectError) {
            console.error(
              `Error fetching sub-projects for project ID ${projectId}:`,
              subProjectError
            );
            failureCount++;
            failedLogs.push({
              date: worklogDate,
              pep: pepValue,
              reason: "API error while fetching sub-projects.",
            });
            continue;
          }
        }

        const date = new Date(worklog.started);
        const formattedDate = `${String(date.getDate()).padStart(
          2,
          "0"
        )}/${String(date.getMonth() + 1).padStart(
          2,
          "0"
        )}/${date.getFullYear()}`;

        const totalMinutes = Math.floor(worklog.timeSpentSeconds / 60);
        const hours = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
        const minutes = String(totalMinutes % 60).padStart(2, "0");
        const formattedHours = `${hours}:${minutes}`;

        const payload = {
          id: "",
          fromDate: formattedDate,
          toDate: formattedDate,
          userId: user.id,
          projectId: projectId,
          category: "PR",
          subProjectId: subProjectId,
          taskId: "", // ignored for now
          description: worklog.comment,
          internalRef: "",
          situationId: "6", // location: in office
          hours: formattedHours,
        };

        try {
          statusDiv.textContent = `Syncing ${
            index + 1
          }/${total}... (Project: ${pepValue})`;
          await postImputation(payload, token);
          successCount++;
        } catch (imputationError) {
          console.error(imputationError);
          failureCount++;
          failedLogs.push({
            date: worklogDate,
            pep: pepValue,
            reason: `Imputation API call failed: ${imputationError.message}`,
          });
        }
      }

      const summaryMessage = `Sync complete. Success: ${successCount}, Failed: ${failureCount}.`;

      if (failedLogs.length > 0) {
        let failureHtml =
          '<div style="margin-top: 10px; font-style: normal; text-align: left;"><strong>Failed Syncs (Page Not Reloaded):</strong><ul style="padding-left: 20px; margin-top: 5px;">';
        for (const log of failedLogs) {
          const pep = log.pep.replace(/</g, "&lt;").replace(/>/g, "&gt;");
          const reason = log.reason.replace(/</g, "&lt;").replace(/>/g, "&gt;");
          failureHtml += `<li style="margin-bottom: 5px;"><strong>${log.date}</strong> (${pep}):<br>${reason}</li>`;
        }
        failureHtml += "</ul></div>";
        statusDiv.innerHTML = summaryMessage + failureHtml;
        console.warn("--- Failed Syncs ---", failedLogs);
      } else {
        statusDiv.textContent =
          summaryMessage + " All worklogs synced. Reloading page...";
        setTimeout(async () => {
          const [tab] = await chrome.tabs.query({
            active: true,
            url: "https://intranetnew.seidor.com/rapports/imputation-hours*",
          });
          if (tab) {
            chrome.tabs.reload(tab.id);
          }
        }, 500); // 0.5 second delay
      }
    } catch (error) {
      console.error("Sync failed:", error);
      statusDiv.textContent = `Error: ${error.message}`;
    }
  }
});
