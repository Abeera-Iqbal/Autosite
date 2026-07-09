// POST /api/trigger
// Body: { reportType: "bk" | "dennys", year, month, dateFromDay?, dateToDay? }
// Dispatches the matching GitHub Actions workflow with a generated run_id,
// then polls the runs list briefly to resolve the actual run id (the
// GitHub dispatch endpoint doesn't return one directly).
//
// Required env vars (set these in your Vercel project settings):
//   GH_OWNER   - GitHub username/org that owns the repo
//   GH_REPO    - repo name
//   GH_PAT     - a fine-grained or classic PAT with "repo" + "workflow" scope
//   GH_BRANCH  - branch to run the workflow from (default "main")

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { reportType, year, month, dateFromDay, dateToDay } = req.body || {};

  if (!reportType || !year || !month) {
    return res.status(400).json({ error: "reportType, year and month are required" });
  }
  if (!["bk", "dennys"].includes(reportType)) {
    return res.status(400).json({ error: "reportType must be 'bk' or 'dennys'" });
  }

  const owner = process.env.GH_OWNER;
  const repo = process.env.GH_REPO;
  const token = process.env.GH_PAT;
  const branch = process.env.GH_BRANCH || "main";

  if (!owner || !repo || !token) {
    return res.status(500).json({ error: "Server is missing GH_OWNER / GH_REPO / GH_PAT env vars" });
  }

  const workflowFile = reportType === "dennys" ? "dennys-report.yml" : "bk-report.yml";
  const runId = `${reportType}-${year}${String(month).padStart(2, "0")}-${Date.now()}`;

  const inputs =
    reportType === "dennys"
      ? {
          year: String(year),
          month: String(month),
          date_from_day: dateFromDay ? String(dateFromDay) : "1",
          date_to_day: dateToDay ? String(dateToDay) : "",
          run_id: runId,
        }
      : {
          year: String(year),
          month: String(month),
          run_id: runId,
        };

  const ghHeaders = {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "Content-Type": "application/json",
  };

  const dispatchedAt = new Date();

  const dispatchResp = await fetch(
    `https://api.github.com/repos/${owner}/${repo}/actions/workflows/${workflowFile}/dispatches`,
    {
      method: "POST",
      headers: ghHeaders,
      body: JSON.stringify({ ref: branch, inputs }),
    }
  );

  if (!dispatchResp.ok) {
    const detail = await dispatchResp.text();
    return res.status(502).json({ error: "Failed to dispatch workflow", detail });
  }

  // GitHub's dispatch endpoint returns 204 with no run id, so poll the runs
  // list for the newest run created after we dispatched.
  let workflowRunId = null;

  for (let i = 0; i < 8; i++) {
    await sleep(1500);

    const runsResp = await fetch(
      `https://api.github.com/repos/${owner}/${repo}/actions/workflows/${workflowFile}/runs?event=workflow_dispatch&per_page=5`,
      { headers: ghHeaders }
    );

    if (!runsResp.ok) continue;

    const runsData = await runsResp.json();
    const match = (runsData.workflow_runs || []).find(
      (run) => new Date(run.created_at) >= dispatchedAt
    );

    if (match) {
      workflowRunId = match.id;
      break;
    }
  }

  return res.status(200).json({ runId, workflowRunId });
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
