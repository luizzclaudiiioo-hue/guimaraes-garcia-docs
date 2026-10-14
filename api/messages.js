import { exigirSessao } from './_auth.js';

// O servidor define o modelo e o limite de resposta; o navegador envia só as mensagens.
const MODELO = 'claude-sonnet-4-5';
const MAX_TOKENS = 1000;
const MAX_CARACTERES = 60000;

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  if (!exigirSessao(req, res)) return;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'API key not configured' });

  const { messages, max_tokens } = req.body || {};
  if (!Array.isArray(messages) || messages.length === 0 || messages.length > 4) {
    return res.status(400).json({ error: 'Mensagens inválidas.' });
  }
  if (JSON.stringify(messages).length > MAX_CARACTERES) {
    return res.status(400).json({ error: 'Texto muito longo.' });
  }

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: MODELO,
        max_tokens: Math.min(Number(max_tokens) || MAX_TOKENS, MAX_TOKENS),
        messages,
      }),
    });

    const data = await response.json();
    return res.status(response.status).json(data);
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
