// Revisão de ortografia e gramática com IA (Claude) para o editor de Texto.
// Recebe os parágrafos do documento e devolve só os erros encontrados, cada um com o trecho
// exato do original, a correção e o motivo. O navegador marca as sugestões como
// alterações controladas (autor "Revisão IA"), que a pessoa aceita ou rejeita.
import Anthropic from '@anthropic-ai/sdk';
import { exigirSessao } from './_auth.js';

const MODELO = 'claude-opus-5-5';
const MAX_CARACTERES = 20000; // por chamada (o navegador divide documentos maiores)
const MAX_PARAGRAFOS = 300;

const INSTRUCOES = `Você é revisor de textos jurídicos em português do Brasil (norma culta e Acordo Ortográfico vigente).
Revise os parágrafos recebidos apontando apenas erros de: ortografia, acentuação, hífen, crase, concordância verbal e nominal, regência, colocação pronominal, pontuação que prejudique a leitura, e palavras repetidas por engano.

Não aponte nem altere:
- escolhas de estilo, vocabulário técnico-jurídico, expressões latinas ou fórmulas usuais de petição;
- nomes próprios, números de processos, datas, valores e referências a leis ou artigos;
- trechos entre aspas que citam lei, jurisprudência ou doutrina (são transcrições);
- o sentido do texto. Na dúvida, não aponte.

Para cada erro, informe:
- "id": o id do parágrafo;
- "trecho": cópia exata, caractere por caractere, do texto original que contém o erro, com 1 a 6 palavras e contexto suficiente para aparecer uma única vez no parágrafo;
- "correcao": o mesmo trecho já corrigido;
- "motivo": explicação curta, com até 12 palavras.
Se não houver erros, devolva a lista vazia.`;

const ESQUEMA = {
  type: 'object',
  properties: {
    correcoes: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'integer' },
          trecho: { type: 'string' },
          correcao: { type: 'string' },
          motivo: { type: 'string' },
        },
        required: ['id', 'trecho', 'correcao', 'motivo'],
        additionalProperties: false,
      },
    },
  },
  required: ['correcoes'],
  additionalProperties: false,
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Use POST.' });
  if (!exigirSessao(req, res)) return;
  if (!process.env.ANTHROPIC_API_KEY) return res.status(500).json({ error: 'Chave da API não configurada no servidor.' });

  const { paragrafos } = req.body || {};
  if (!Array.isArray(paragrafos) || !paragrafos.length || paragrafos.length > MAX_PARAGRAFOS
    || paragrafos.some((p) => !Number.isInteger(p?.id) || typeof p?.texto !== 'string')) {
    return res.status(400).json({ error: 'Parágrafos inválidos.' });
  }
  const total = paragrafos.reduce((a, p) => a + p.texto.length, 0);
  if (total > MAX_CARACTERES) return res.status(400).json({ error: 'Trecho grande demais para uma revisão.' });

  try {
    const client = new Anthropic();
    const resposta = await client.beta.messages.create({
      model: MODELO,
      max_tokens: 16000,
      // Se o modelo recusar o pedido (filtro de segurança), outro modelo da Anthropic assume.
      betas: ['server-side-fallback-2026-07-01'],
      fallbacks: 'default',
      system: INSTRUCOES,
      output_config: { effort: 'medium', format: { type: 'json_schema', schema: ESQUEMA } },
      messages: [{ role: 'user', content: JSON.stringify({ paragrafos }) }],
    });
    if (resposta.stop_reason === 'refusal') return res.status(422).json({ error: 'A IA não revisou este trecho.' });
    if (resposta.stop_reason === 'max_tokens') return res.status(422).json({ error: 'Correções demais neste trecho; selecione uma parte menor.' });
    const texto = resposta.content.find((b) => b.type === 'text')?.text || '{"correcoes":[]}';
    const { correcoes } = JSON.parse(texto);
    return res.status(200).json({ correcoes, uso: { entrada: resposta.usage.input_tokens, saida: resposta.usage.output_tokens } });
  } catch (e) {
    if (e instanceof Anthropic.RateLimitError) return res.status(429).json({ error: 'Muitas revisões ao mesmo tempo; tente de novo em instantes.' });
    if (e instanceof Anthropic.APIError) return res.status(502).json({ error: `Serviço de IA indisponível (${e.status ?? 'erro'}).` });
    return res.status(500).json({ error: 'Não foi possível revisar: ' + (e?.message || e) });
  }
}
