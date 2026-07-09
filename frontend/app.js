// ── Config ──────────────────────────────────────────────────────────────────
// Same-origin serverless functions (see /api). Change if you host the API
// functions somewhere other than where this static site is served from.
const API_BASE = "/api";

// Public Test Records API that the automation scripts already upload to.
const FILES_API = "https://distapi.cybussolutions.com/api/v1/test-records/files";

const STATUS_POLL_MS = 5000;
const FILES_POLL_MS = 4000;
const FILES_POLL_ATTEMPTS = 20; // ~80s of polling once the workflow says "completed"

// ── Elements ────────────────────────────────────────────────────────────────
const sourceToggle = document.getElementById("sourceToggle");
const periodInput = document.getElementById("period");
const dennysDateRange = document.getElementById("dennysDateRange");
const dayFromInput = document.getElementById("dayFrom");
const dayToInput = document.getElementById("dayTo");
const runBtn = document.getElementById("runBtn");

const statusBox = document.getElementById("status");
const statusDot = document.getElementById("statusDot");
const statusText = document.getElementById("statusText");
const logsLink = document.getElementById("logsLink");

const resultsBox = document.getElementById("results");
const resultsList = document.getElementById("resultsList");

const errorBox = document.getElementById("errorBox");

let reportType = "bk";

// ── Init ────────────────────────────────────────────────────────────────────
(function init() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  periodInput.value = `${yyyy}-${mm}`;
})();

sourceToggle.addEventListener("click", (e) => {
  const btn = e.target.closest(".toggle");
  if (!btn) return;
  reportType = btn.dataset.value;
  [...sourceToggle.children].forEach((c) => c.classList.toggle("active", c === btn));
  dennysDateRange.hidden = reportType !== "dennys";
});

runBtn.addEventListener("click", runReport);

// ── Main flow ───────────────────────────────────────────────────────────────
async function runReport() {
  resetUI();

  const [year, month] = periodInput.value.split("-").map(Number);
  if (!year || !month) {
    showError("Pick a month first.");
    return;
  }

  const payload = {
    reportType,
    year,
    month,
    dateFromDay: reportType === "dennys" ? Number(dayFromInput.value) || 1 : undefined,
    dateToDay: reportType === "dennys" && dayToInput.value ? Number(dayToInput.value) : undefined,
  };

  runBtn.disabled = true;
  runBtn.textContent = "Starting…";
  statusBox.hidden = false;
  setStatus("queued", "Triggering GitHub Action…");

  try {
    const trigger = await postJSON(`${API_BASE}/trigger`, payload);

    if (!trigger.workflowRunId) {
      setStatus("running", "Workflow dispatched — waiting for GitHub to register the run…");
      // one more short poll attempt in case /api/trigger returned before resolving it
      await sleep(4000);
    }

    const workflowRunId = trigger.workflowRunId;
    runBtn.textContent = "Running…";

    if (workflowRunId) {
      await pollWorkflowStatus(workflowRunId);
    }

    setStatus("running", "Workflow finished — looking for your files…");
    const files = await pollForFiles(trigger.runId);

    if (files.length === 0) {
      setStatus("failure", "Workflow completed but no matching files were found yet.");
      showError(
        "The run finished but the report files haven't shown up on the Test Records API yet. " +
        "Try clicking Run report again in a minute, or check the run logs."
      );
    } else {
      setStatus("success", "Done — your files are ready.");
      showResults(files);
    }
  } catch (err) {
    setStatus("failure", "Something went wrong.");
    showError(err.message || String(err));
  } finally {
    runBtn.disabled = false;
    runBtn.textContent = "Run report";
  }
}

async function pollWorkflowStatus(workflowRunId) {
  while (true) {
    const data = await getJSON(`${API_BASE}/status?workflowRunId=${workflowRunId}`);

    if (data.htmlUrl) {
      logsLink.href = data.htmlUrl;
      logsLink.hidden = false;
    }

    if (data.status === "completed") {
      if (data.conclusion === "success") {
        setStatus("success", "GitHub Action completed successfully.");
      } else {
        setStatus("failure", `GitHub Action finished with: ${data.conclusion}`);
        throw new Error(`Workflow run finished with conclusion "${data.conclusion}". Check the logs.`);
      }
      return;
    }

    setStatus("running", `GitHub Action is ${data.status.replace("_", " ")}…`);
    await sleep(STATUS_POLL_MS);
  }
}

async function pollForFiles(runId) {
  for (let i = 0; i < FILES_POLL_ATTEMPTS; i++) {
    const resp = await fetch(FILES_API);
    if (resp.ok) {
      const body = await resp.json();
      const all = body.data || [];
      const matches = all.filter((f) => f.fileName && f.fileName.includes(runId));
      if (matches.length > 0) return matches;
    }
    await sleep(FILES_POLL_MS);
  }
  return [];
}

// ── UI helpers ──────────────────────────────────────────────────────────────
function resetUI() {
  errorBox.hidden = true;
  errorBox.textContent = "";
  resultsBox.hidden = true;
  resultsList.innerHTML = "";
  logsLink.hidden = true;
  statusBox.hidden = true;
}

function setStatus(kind, text) {
  statusDot.className = `dot ${kind === "queued" ? "" : kind}`;
  statusText.textContent = text;
}

function showResults(files) {
  resultsBox.hidden = false;
  resultsList.innerHTML = "";
  files.forEach((f) => {
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = `https://distapi.cybussolutions.com${f.downloadUrl}`;
    a.target = "_blank";
    a.rel = "noopener";
    a.textContent = `⬇ ${f.fileName}`;
    li.appendChild(a);
    resultsList.appendChild(li);
  });
}

function showError(msg) {
  errorBox.hidden = false;
  errorBox.textContent = msg;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function postJSON(url, body) {
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.error || `Request failed (${resp.status})`);
  return data;
}

async function getJSON(url) {
  const resp = await fetch(url);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.error || `Request failed (${resp.status})`);
  return data;
}
