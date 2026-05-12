import { neon } from "@neondatabase/serverless";

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  try {
    if (req.method === "GET") {
      const rows = await sql`SELECT data, updated_at FROM app_state WHERE id = 'procurements'`;
      if (rows.length === 0) return res.status(200).json({ data: null, updatedAt: null });
      return res.status(200).json({ data: rows[0].data, updatedAt: rows[0].updated_at });
    }
    if (req.method === "PUT") {
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      if (!body || typeof body !== "object") return res.status(400).json({ error: "Body must be an object" });
      await sql`
        INSERT INTO app_state (id, data, updated_at)
        VALUES ('procurements', ${body}::jsonb, NOW())
        ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data, updated_at = NOW()
      `;
      return res.status(200).json({ ok: true });
    }
    res.setHeader("Allow", "GET, PUT");
    return res.status(405).json({ error: "Method not allowed" });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
