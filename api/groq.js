// Proxy Groq pour hébergement (Vercel Serverless) — évite CORS navigateur
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { apiKey, model, messages, temperature, max_tokens, response_format } = req.body || {};
  if (!apiKey) return res.status(400).json({ error: { message: 'apiKey manquante' } });

  try {
    const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: model || 'openai/gpt-oss-20b',
        messages,
        temperature: temperature ?? 0.3,
        max_tokens: max_tokens ?? 2048,
        ...(response_format ? { response_format } : {}),
      }),
    });

    const data = await groqRes.json();
    return res.status(groqRes.status).json(data);
  } catch (e) {
    return res.status(500).json({ error: { message: e.message } });
  }
}
