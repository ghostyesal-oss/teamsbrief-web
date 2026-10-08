// Proxy Groq pour Netlify Functions — évite CORS navigateur
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: CORS, body: '' };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: CORS, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  let body;
  try {
    body = JSON.parse(event.body || '{}');
  } catch {
    return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: { message: 'JSON invalide' } }) };
  }

  const { apiKey, model, messages, temperature, max_tokens, response_format } = body;
  if (!apiKey) {
    return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: { message: 'apiKey manquante' } }) };
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
    return {
      statusCode: groqRes.status,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    };
  } catch (e) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ error: { message: e.message } }),
    };
  }
};
