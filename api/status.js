// GET /api/status?workflowRunId=123456
// Returns { status, conclusion, htmlUrl } for a GitHub Actions run.

export default async function handler(req, res) {
  const { workflowRunId } = req.query;

  if (!workflowRunId) {
    return res.status(400).json({ error: "workflowRunId query param is required" });
  }

  const owner = process.env.GH_OWNER;
  const repo = process.env.GH_REPO;
  const token = process.env.GH_PAT;

  if (!owner || !repo || !token) {
    return res.status(500).json({ error: "Server is missing GH_OWNER / GH_REPO / GH_PAT env vars" });
  }

  const resp = await fetch(
    `https://api.github.com/repos/${owner}/${repo}/actions/runs/${workflowRunId}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json",
      },
    }
  );

  if (!resp.ok) {
    const detail = await resp.text();
    return res.status(502).json({ error: "Failed to fetch run status", detail });
  }

  const data = await resp.json();

  return res.status(200).json({
    status: data.status,          // queued | in_progress | completed
    conclusion: data.conclusion,  // success | failure | cancelled | null
    htmlUrl: data.html_url,
  });
}
