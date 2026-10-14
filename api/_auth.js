// Autenticação das rotas /api: senha conferida no servidor e sessão em cookie assinado (HMAC).
// Arquivos de /api que começam com "_" não viram rotas na Vercel.
import crypto from 'crypto';

const COOKIE = 'sessao';
const DURACAO_SEG = 12 * 60 * 60;

function senhas() {
  return [
    process.env.SENHA_1 || process.env.VITE_SENHA_1,
    process.env.SENHA_2 || process.env.VITE_SENHA_2,
  ].filter(Boolean);
}

function segredo() {
  if (process.env.SEGREDO_SESSAO) return process.env.SEGREDO_SESSAO;
  const base = [process.env.ANTHROPIC_API_KEY, ...senhas()].filter(Boolean).join('|');
  if (!base) throw new Error('Senhas de acesso não configuradas.');
  return crypto.createHash('sha256').update('sessao|' + base).digest('hex');
}

function assinar(valor) {
  return crypto.createHmac('sha256', segredo()).update(valor).digest('hex');
}

function iguais(a, b) {
  const ha = crypto.createHash('sha256').update(String(a)).digest();
  const hb = crypto.createHash('sha256').update(String(b)).digest();
  return crypto.timingSafeEqual(ha, hb);
}

export function senhaCorreta(senha) {
  return !!senha && senhas().some((s) => iguais(senha, s));
}

export function criarSessao(res) {
  const expira = String(Math.floor(Date.now() / 1000) + DURACAO_SEG);
  const token = `${expira}.${assinar(expira)}`;
  res.setHeader('Set-Cookie', `${COOKIE}=${token}; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=${DURACAO_SEG}`);
}

export function sessaoValida(req) {
  const cookies = String(req.headers.cookie || '').split(';').map((c) => c.trim());
  const token = cookies.find((c) => c.startsWith(COOKIE + '='))?.slice(COOKIE.length + 1);
  if (!token) return false;
  const [expira, assinatura] = token.split('.');
  if (!expira || !assinatura || Number(expira) < Date.now() / 1000) return false;
  try {
    return iguais(assinatura, assinar(expira));
  } catch {
    return false;
  }
}

// Responde 401 e devolve false quando não há sessão válida.
export function exigirSessao(req, res) {
  if (sessaoValida(req)) return true;
  res.status(401).json({ error: 'Sessão expirada. Recarregue a página e entre novamente.' });
  return false;
}
