// Proxy Groq pour Cloudflare Pages Functions — évite CORS navigateur
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export async function onRequest(context) {
  const { request } = context;

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: CORS });
  }
  if (request.method !== 'POST') {
    return Response.json({ error: 'Method not allowed' }, { status: 405, headers: CORS });
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return Response.json({ error: { message: 'JSON invalide' } }, { status: 400, headers: CORS });
  }

  const { apiKey, model, messages, temperature, max_tokens, response_format } = body;
  if (!apiKey) {
    return Response.json({ error: { message: 'apiKey manquante' } }, { status: 400, headers: CORS });
  }

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
    return Response.json(data, { status: groqRes.status, headers: CORS });
  } catch (e) {
    return Response.json({ error: { message: e.message } }, { status: 500, headers: CORS });
  }
}
