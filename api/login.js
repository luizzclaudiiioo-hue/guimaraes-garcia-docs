import { senhaCorreta, criarSessao, sessaoValida } from './_auth.js';

export default async function handler(req, res) {
  // GET: informa se a sessão atual ainda vale (ao recarregar a página)
  if (req.method === 'GET') return res.status(sessaoValida(req) ? 200 : 401).json({});
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    if (!senhaCorreta(req.body?.senha)) {
      // Atraso para dificultar tentativas em massa
      await new Promise((r) => setTimeout(r, 1000));
      return res.status(401).json({ error: 'Senha incorreta.' });
    }
    criarSessao(res);
    return res.status(200).json({});
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
