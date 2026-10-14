// Consulta as publicações (intimações) do DJEN — Comunica PJe/CNJ — pela OAB/UF.
const DJEN_URL = 'https://comunicaapi.pje.jus.br/api/v1/comunicacao';
const ITENS_POR_PAGINA = 100;
const MAX_PAGINAS = 10;

const dataValida = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s || '');

export default async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const oab = String(req.query.oab || '').replace(/\D/g, '');
  const uf = String(req.query.uf || '').toUpperCase();
  const { inicio, fim } = req.query;

  if (!oab) return res.status(400).json({ error: 'Informe o número da OAB.' });
  if (!/^[A-Z]{2}$/.test(uf)) return res.status(400).json({ error: 'Informe a UF da OAB (ex: SP).' });
  if (!dataValida(inicio) || !dataValida(fim)) return res.status(400).json({ error: 'Período inválido.' });

  try {
    const itens = [];
    let total = 0;
    for (let pagina = 1; pagina <= MAX_PAGINAS; pagina++) {
      const params = new URLSearchParams({
        numeroOab: oab,
        ufOab: uf,
        dataDisponibilizacaoInicio: inicio,
        dataDisponibilizacaoFim: fim,
        pagina: String(pagina),
        itensPorPagina: String(ITENS_POR_PAGINA),
      });
      const resp = await fetch(`${DJEN_URL}?${params}`, {
        headers: { Accept: 'application/json' },
        signal: AbortSignal.timeout(20000),
      });
      if (!resp.ok) {
        const corpo = await resp.text().catch(() => '');
        return res.status(502).json({ error: `O DJEN respondeu com erro ${resp.status}.`, detalhe: corpo.slice(0, 300) });
      }
      const data = await resp.json();
      const lote = data.items || [];
      total = data.count ?? lote.length;
      itens.push(...lote);
      if (lote.length < ITENS_POR_PAGINA || itens.length >= total) break;
    }

    const publicacoes = itens.map((p) => ({
      id: p.id ?? p.hash,
      data: p.data_disponibilizacao || p.datadisponibilizacao || '',
      tribunal: p.siglaTribunal || '',
      orgao: p.nomeOrgao || '',
      processo: p.numeroprocessocommascara || p.numero_processo || '',
      classe: p.nomeClasse || '',
      tipo: p.tipoComunicacao || p.tipoDocumento || '',
      texto: p.texto || '',
      link: p.link || '',
      partes: (p.destinatarios || []).map((d) => d.nome).filter(Boolean),
    }));
    publicacoes.sort((a, b) => String(b.data).localeCompare(String(a.data)));

    return res.status(200).json({ total, publicacoes });
  } catch (error) {
    return res.status(502).json({ error: 'Não foi possível consultar o DJEN: ' + error.message });
  }
}
