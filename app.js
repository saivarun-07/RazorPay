const thread = document.querySelector("#thread");
const input = document.querySelector("#questionInput");
const composer = document.querySelector("#composer");
const toast = document.querySelector("#toast");
const sourceInput = document.querySelector("#sourceInput");
const chatView = document.querySelector(".content-grid");
const sourcesView = document.querySelector("#sourcesView");
const activityView = document.querySelector("#activityView");
const settingsView = document.querySelector("#settingsView");
const navItems = document.querySelectorAll(".main-nav .nav-item");
const workbookList = document.querySelector("#workbookList");
const sidebarSourceList = document.querySelector("#sidebarSourceList");
const uploadCard = document.querySelector("#uploadCard");
const activityList = document.querySelector(".activity-list");
const apiBase =
  window.location.hostname === "localhost" && window.location.port !== "8000"
    ? "http://localhost:8000"
    : "";
let workbooks = [];
let activity = [];
let sessionSummary = "Q2 software spend and its quarter-over-quarter change.";

function setView(view) {
  const normalizedView =
    view === "sources" || view === "activity" || view === "settings"
      ? view
      : "chat";
  chatView.style.display = normalizedView === "chat" ? "grid" : "none";
  sourcesView.style.display = normalizedView === "sources" ? "block" : "none";
  activityView.style.display = normalizedView === "activity" ? "block" : "none";
  settingsView.style.display = normalizedView === "settings" ? "block" : "none";
  sourcesView.classList.toggle("active", normalizedView === "sources");
  activityView.classList.toggle("active", normalizedView === "activity");
  settingsView.classList.toggle("active", normalizedView === "settings");
  navItems.forEach((item) =>
    item.classList.toggle(
      "active",
      item.getAttribute("href") === `#${normalizedView}` ||
        (normalizedView === "chat" && item.getAttribute("href") === "#chat"),
    ),
  );
  document.querySelector(".breadcrumbs span:first-child").textContent =
    normalizedView === "chat"
      ? "Ask Ledgerly"
      : normalizedView === "sources"
        ? "Data sources"
        : normalizedView === "activity"
          ? "Activity"
          : "Settings";
}

function renderActivity() {
  if (!activityList) return;
  activityList.innerHTML = activity.length
    ? activity
        .map(
          (item) =>
            `<div class="activity-row"><span class="activity-icon green-icon">✓</span><div><strong>${escapeHtml(item.kind)}</strong><p>${escapeHtml(item.detail)}</p></div><time>Just now</time></div>`,
        )
        .join("")
    : '<div class="activity-row"><span class="activity-icon">•</span><div><strong>No activity yet</strong><p>Workbook imports and verified answers will appear here.</p></div><time>Now</time></div>';
}

async function loadState() {
  try {
    const response = await fetch(`${apiBase}/api/state`);
    if (!response.ok) throw new Error("State unavailable");
    const state = await response.json();
    workbooks = state.workbooks || [];
    activity = state.activity || [];
    renderWorkbooks();
    renderActivity();
  } catch {
    showToast("Connect to server.py to load live data");
  }
}

window.addEventListener("hashchange", () =>
  setView(window.location.hash.slice(1)),
);
setView(window.location.hash.slice(1));

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(
    () => toast.classList.remove("show"),
    2400,
  );
}

function renderWorkbooks() {
  workbookList
    .querySelectorAll(".workbook-card")
    .forEach((card) => card.remove());
  workbooks.forEach((workbook) => {
    const card = document.createElement("div");
    card.className = "workbook-card";
    card.dataset.workbookName = workbook.name;
    card.dataset.tables = JSON.stringify(workbook.tables || []);
    card.innerHTML = `<div class="workbook-card-top"><span class="file-icon">▦</span><span class="source-status"></span></div><h2></h2><p>${workbook.status} · ${workbook.sheets} · ${workbook.rows}</p><div class="sheet-list"><span>Imported workbook</span></div><div class="workbook-footer"><span>${workbook.status}</span><button class="copy-button remove-workbook">Remove</button></div>`;
    card.querySelector("h2").textContent = workbook.name;
    workbookList.insertBefore(card, uploadCard);
  });
  sidebarSourceList.innerHTML = workbooks.length
    ? workbooks
        .map(
          (workbook) =>
            `<span class="db-icon">▣</span><div><strong>${workbook.name}</strong><small>Excel · ${workbook.sheets}</small></div><span class="source-status"></span>`,
        )
        .join("")
    : '<span class="db-icon">＋</span><div><strong>No workbook connected</strong><small>Upload an Excel source to begin</small></div>';
}

function escapeHtml(value) {
  return String(value).replace(
    /[&<>]/g,
    (character) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
      })[character],
  );
}

function renderMarkdown(text) {
  const raw = marked.parse(String(text ?? ""), { breaks: true });
  return DOMPurify.sanitize(raw);
}

function getAssistantMarkup() {
  return '<div class="message-avatar assistant-avatar">✦</div><div class="message-body"><div class="message-meta"><strong>Ledgerly</strong><span>Just now · <span class="verified">✓ Fetching workbook data</span></span></div><p>Fetching the relevant rows from your imported workbook...</p><div class="evidence-card"><div class="evidence-header"><div><span class="evidence-icon">▤</span><strong>Fetching data</strong></div><span class="row-count">In progress</span></div><div class="evidence-footer"><span>Source: <strong>Excel workbook</strong></span><span class="copy-button">Reading rows</span></div></div><div class="guardrail"><span>◈</span> No estimates used. The assistant will not invent financial values.</div></div>';
}

function appendAiResponse(message, response) {
  const evidence = response.evidence?.[0];
  const evidenceResult = evidence?.result || evidence;
  const evidenceQuery =
    evidence?.arguments?.query || evidenceResult?.query || "";
  const assistantMessage = document.createElement("div");
  assistantMessage.className = "message assistant-message";
  assistantMessage.innerHTML = `<div class="message-avatar assistant-avatar">✦</div><div class="message-body"><div class="message-meta"><strong>Ledgerly</strong><span>Just now · <span class="verified">✓ AI answer from workbook</span></span></div><div class="answer-markdown">${renderMarkdown(response.answer || "No verified answer returned.")}</div>${evidence ? `<div class="evidence-card"><div class="evidence-header"><div><span class="evidence-icon">▤</span><strong>Answer evidence</strong></div><span class="row-count">${evidenceResult?.row_count ?? 0} matching rows</span></div><button class="sql-toggle response-sql-toggle"><span>⌘</span> View SQL used <span class="toggle-arrow">⌄</span></button><pre class="sql-block response-sql-block"><code>${escapeHtml(evidenceQuery)}</code></pre><div class="evidence-footer"><span>Source: <strong>imported Excel sheet</strong></span><button class="copy-button response-copy">Copy SQL</button></div></div>` : ""}<div class="guardrail"><span>◈</span> No estimates used. Every figure is traceable to imported workbook rows.</div></div>`;
  thread.appendChild(assistantMessage);
  const toggle = assistantMessage.querySelector(".response-sql-toggle");
  if (toggle) {
    const sqlBlock = assistantMessage.querySelector(".response-sql-block");
    toggle.addEventListener("click", () => {
      sqlBlock.classList.toggle("visible");
      toggle.classList.toggle("open");
    });
    assistantMessage
      .querySelector(".response-copy")
      .addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(evidenceQuery);
          showToast("SQL copied to clipboard");
        } catch {
          showToast("Select the SQL block to copy it");
        }
      });
  }
}

renderWorkbooks();
renderActivity();
loadState();

async function addMessage(question) {
  const userMessage = document.createElement("div");
  userMessage.className = "message user-message";
  userMessage.innerHTML = `<div class="message-avatar user-avatar">AK</div><div class="message-body"><div class="message-meta"><strong>You</strong><span>Just now</span></div><p>${question.replace(/[&<>]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[char])}</p></div>`;
  thread.appendChild(userMessage);

  const assistantMessage = document.createElement("div");
  assistantMessage.className = "message assistant-message pending-message";
  assistantMessage.innerHTML = getAssistantMarkup(question);
  thread.appendChild(assistantMessage);
  const projectionToggle = assistantMessage.querySelector(
    ".projection-sql-toggle",
  );
  const projectionSql = assistantMessage.querySelector(".projection-sql-block");
  if (projectionToggle) {
    projectionToggle.addEventListener("click", () => {
      projectionSql.classList.toggle("visible");
      projectionToggle.classList.toggle("open");
    });
    assistantMessage
      .querySelector(".projection-copy")
      .addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(projectionSql.textContent.trim());
          showToast("SQL copied to clipboard");
        } catch {
          showToast("Select the SQL block to copy it");
        }
      });
  }
  try {
    const response = await fetch(`${apiBase}/api/ask`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question, session_summary: sessionSummary }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "AI request failed");
    assistantMessage.remove();
    appendAiResponse(question, result);
    sessionSummary = `Current question: ${question}. Answer grounded in ${result.evidence?.[0]?.row_count ?? 0} workbook rows.`;
    document.querySelector(".context-summary p").textContent = sessionSummary;
    activity.unshift({ kind: "Answer generated", detail: question });
    renderActivity();
  } catch (error) {
    assistantMessage.querySelector("p").textContent =
      `AI is not connected: ${error.message.includes("403") ? "Groq rejected the API key" : error.message}. Check GROQ_API_KEY in .env and restart server.py.`;
    assistantMessage.querySelector(".verified").textContent =
      "✕ AI request failed";
    showToast("AI request failed");
  }
  input.value = "";
  input.style.height = "auto";
  showToast("Question added to this session summary");
}

composer.addEventListener("submit", (event) => {
  event.preventDefault();
  const question = input.value.trim();
  if (question) addMessage(question);
});

document.querySelectorAll(".suggestion").forEach((button) => {
  button.addEventListener("click", () => {
    input.value = button.textContent.replace("→", "").trim();
    input.focus();
  });
});

document.querySelector("#clearSession").addEventListener("click", () => {
  thread.innerHTML = "";
  document.querySelector(".context-summary p").innerHTML =
    "<strong>New session</strong> — no previous conversation context is carried forward.";
  sessionSummary = "New session with no previous conversation context.";
  showToast("Session context cleared");
});

document.querySelector("#newQuestion").addEventListener("click", () => {
  setView("chat");
  document.querySelector(".context-summary p").innerHTML =
    "This session is about <strong>a new financial question</strong>.";
  sessionSummary = "A new financial question.";
  input.value = "";
  input.focus();
  showToast("Started a new question");
});

document
  .querySelector("#manageSources")
  .addEventListener("click", () => sourceInput.click());
document.querySelector("#addSource").addEventListener("click", () => {
  sourceInput.click();
});
document
  .querySelector("#sourcesUpload")
  .addEventListener("click", () => sourceInput.click());
document
  .querySelector("#uploadCard")
  .addEventListener("click", () => sourceInput.click());

document.querySelector(".workspace-switcher").addEventListener("click", () => {
  showToast("Acme Finance is the active workspace");
});
document.querySelector(".help-button").addEventListener("click", () => {
  showToast("Ask about imported workbook data or general finance concepts");
});
document.querySelector(".icon-button").addEventListener("click", () => {
  document.querySelector(".notification-dot").hidden = true;
  window.location.hash = "activity";
  showToast(
    activity.length
      ? `${activity.length} recent activity item${activity.length === 1 ? "" : "s"}`
      : "No new notifications",
  );
});
document.querySelector(".model-select").addEventListener("click", () => {
  showToast("GPT OSS 20B is the active reasoning model");
});
const groundedSetting = document.querySelector("#groundedSetting");
const auditSetting = document.querySelector("#auditSetting");
groundedSetting.checked = localStorage.getItem("ledgerly-grounded") !== "false";
auditSetting.checked = localStorage.getItem("ledgerly-audit") !== "false";
document.querySelector("#saveSettings").addEventListener("click", () => {
  localStorage.setItem("ledgerly-grounded", groundedSetting.checked);
  localStorage.setItem("ledgerly-audit", auditSetting.checked);
  showToast("Preferences saved");
});
document.querySelector(".more").addEventListener("click", () => {
  showToast("Workbook connection is healthy");
});
workbookList.addEventListener("click", (event) => {
  const removeButton = event.target.closest(".remove-workbook");
  if (!removeButton) return;
  const card = removeButton.closest(".workbook-card");
  const name = card.dataset.workbookName;
  removeButton.disabled = true;
  fetch(`${apiBase}/api/workbooks`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      tables: card.dataset.tables ? JSON.parse(card.dataset.tables) : [],
    }),
  })
    .then(async (response) => {
      const result = await response.json();
      if (!response.ok)
        throw new Error(result.error || "Workbook removal failed");
      workbooks = workbooks.filter((workbook) => workbook.name !== name);
      renderWorkbooks();
      activity.unshift({ kind: "Workbook removed", detail: name });
      renderActivity();
      showToast(`${name} removed`);
    })
    .catch((error) => {
      removeButton.disabled = false;
      showToast(error.message);
    });
});
sourceInput.addEventListener("change", () => {
  const [workbook] = sourceInput.files;
  if (!workbook) return;
  if (workbooks.some((item) => item.name === workbook.name)) {
    showToast(`${workbook.name} is already connected`);
    sourceInput.value = "";
    return;
  }
  workbooks.push({
    name: workbook.name,
    sheets: "Pending import",
    rows: "Awaiting import",
    status: "Selected just now",
  });
  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const response = await fetch(`${apiBase}/api/workbooks`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: workbook.name,
          content: reader.result.split(",")[1],
        }),
      });
      if (!response.ok) throw new Error("Workbook import failed");
      const imported = await response.json();
      const current = workbooks.find((item) => item.name === workbook.name);
      if (!current) return;
      current.tables = imported.sheets.map((sheet) => sheet.table);
      current.sheets = `${imported.sheets.length} sheets`;
      current.rows = `${imported.sheets.reduce((total, sheet) => total + sheet.rows, 0)} rows`;
      current.status = "Imported just now";
      renderWorkbooks();
      activity.unshift({
        kind: "Workbook imported",
        detail: `${workbook.name} · ${imported.sheets.length} sheets`,
      });
      renderActivity();
      showToast(`${workbook.name} imported for AI queries`);
    } catch (error) {
      showToast(error.message);
      workbooks = workbooks.filter((item) => item.name !== workbook.name);
      renderWorkbooks();
    }
  };
  reader.readAsDataURL(workbook);
  renderWorkbooks();
  showToast(`${workbook.name} added`);
  sourceInput.value = "";
});
input.addEventListener("input", () => {
  input.style.height = "auto";
  input.style.height = `${Math.min(input.scrollHeight, 100)}px`;
});
