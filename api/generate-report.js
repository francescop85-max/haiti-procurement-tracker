export const config = { runtime: "edge" };

const MODEL = "claude-sonnet-4-5-20250929";
const ANTHROPIC_VERSION = "2023-06-01";

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function buildPrompt(s) {
  const fmtUsd = (n) => "$" + Number(n || 0).toLocaleString("en-US");

  const cats = (s.topCategories || [])
    .map((c) => `- ${c.category}: ${fmtUsd(c.value)} (${c.count} procurement${c.count === 1 ? "" : "s"})`)
    .join("\n") || "(none)";

  const atRisk = (s.atRisk || [])
    .map((r) => {
      const id = r.lot ? `${r.pr} – ${r.lot}` : r.pr;
      const buf = r.buffer == null ? "no target" : `${r.buffer}d buffer`;
      return `- [${r.riskLevel.toUpperCase()}] ${r.tender} (PR ${id}, ${r.method}, ${fmtUsd(r.estInitial)}) — ${r.riskLabel}, ${buf}. Status: ${r.narrative || "—"}`;
    })
    .join("\n") || "(none)";

  const upcoming = (s.upcomingPOs || [])
    .map((u) => {
      const id = u.lot ? `${u.pr} – ${u.lot}` : u.pr;
      return `- ${u.tender} (PR ${id}, ${u.method}, ${fmtUsd(u.estInitial)}): target ${u.targetPO || "—"}, est PO ${u.estPODate || "—"} (${u.daysOut}d out)`;
    })
    .join("\n") || "(none)";

  return `You are writing the executive summary section of a management report for FAO Haiti's procurement portfolio.

SNAPSHOT DATE: ${s.date}

PORTFOLIO METRICS
- Portfolio total (PR estimate): ${fmtUsd(s.portfolioTotal)}
- Active value: ${fmtUsd(s.activeTotal)}
- Cancelled value: ${fmtUsd(s.cancelledTotal)}
- Active procurements: ${s.countActive}
- Cancelled: ${s.countCancelled}
- Pre-pipeline / not applicable: ${s.countPipeline}
- Average PR→PO variance: ${s.avgVariance}%

RISK PROFILE
- On track: ${s.ok}
- Watch: ${s.watch}
- Critical: ${s.critical}

TOP CATEGORIES BY VALUE
${cats}

PROCUREMENTS REQUIRING ATTENTION
${atRisk}

UPCOMING PO MILESTONES (next 60 days)
${upcoming}

TASK
Write a 2-3 paragraph executive summary suitable for the FAO Haiti Country Director and senior procurement management. Cover, in order:
1. Overall portfolio health and the headline numbers
2. Specific risks that need management attention (call out tenders by name and value)
3. What is expected to land in the coming weeks and any forward-looking considerations

GUIDELINES
- Use specific numbers and tender names from the snapshot. Do not invent data, dates, or facts.
- Professional, concise, action-oriented tone. Plain prose only — no bullet points, no markdown headings, no section labels.
- 250–350 words total.
- Output only the summary text. No preamble, no sign-off.`;
}

export default async function handler(req) {
  if (req.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405);
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return jsonResponse(
      { error: "ANTHROPIC_API_KEY is not configured on the server. Set it in Vercel project settings." },
      500,
    );
  }

  let snapshot;
  try {
    snapshot = await req.json();
  } catch {
    return jsonResponse({ error: "Invalid JSON body" }, 400);
  }

  try {
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": ANTHROPIC_VERSION,
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 1200,
        messages: [{ role: "user", content: buildPrompt(snapshot) }],
      }),
    });

    if (!upstream.ok) {
      const detail = await upstream.text();
      return jsonResponse({ error: `Anthropic API ${upstream.status}: ${detail}` }, 502);
    }

    const data = await upstream.json();
    const narrative = data?.content?.[0]?.text?.trim();
    if (!narrative) {
      return jsonResponse({ error: "Anthropic returned no content" }, 502);
    }

    return jsonResponse({ narrative });
  } catch (err) {
    return jsonResponse({ error: err.message || "Unknown error" }, 500);
  }
}
