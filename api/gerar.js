import JSZip from 'jszip';

const GITHUB_USER = 'luizzclaudiiioo-hue';
const GITHUB_REPO = 'guimaraes-garcia-docs';
const GITHUB_BRANCH = 'main';

async function fetchTemplate(filename) {
  const url = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/${GITHUB_BRANCH}/templates/${filename}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Template not found: ${filename} (${res.status})`);
  return Buffer.from(await res.arrayBuffer());
}

function isRed(rPrXml) {
  return /<w:color[^>]*w:val="FF0000"/.test(rPrXml);
}

function removeRedColor(rPrInner) {
  return rPrInner
    .replace(/<w:color[^/]*\/>/g, '')
    .replace(/<w:color[^>]*>[\s\S]*?<\/w:color>/g, '');
}

function esc(s) {
  return (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function substituirPorIndice(xml, mapa) {
  let redIdx = 0;
  return xml.replace(/(<w:r)([ >])([\s\S]*?)(<\/w:r>)/g, (match, tag, sep, inner, close) => {
    const rPrMatch = inner.match(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/);
    if (rPrMatch && isRed(rPrMatch[0])) {
      const currentIdx = redIdx++;
      const entry = mapa.find(m => m.index === currentIdx);
      const valor = entry !== undefined ? esc(entry.value) : '';
      const newRPr = rPrMatch[1] + removeRedColor(rPrMatch[2]) + rPrMatch[3];
      let newInner = inner.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, newRPr);
      newInner = newInner.replace(/<w:t[^>]*>[\s\S]*?<\/w:t>/, `<w:t xml:space="preserve">${valor}</w:t>`);
      return tag + sep + newInner + close;
    }
    return match;
  });
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { tipo, dados: d, financeiro: fin } = req.body;
    const hoje = new Date();
    const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
    const dia = String(hoje.getDate());
    const mes = meses[hoje.getMonth()];
    const ano = String(hoje.getFullYear());

    const filename = tipo === 'procuracao' ? 'template_procuracao.docx' : 'template_contrato.docx';
    const templateBuf = await fetchTemplate(filename);
    const zip = await JSZip.loadAsync(templateBuf);
    let xml = await zip.file('word/document.xml').async('string');

    if (tipo === 'procuracao') {
      const nome = d.nome.toUpperCase();
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: '' },
        { index: 2,  value: 'brasileiro(a)' },
        { index: 3,  value: '' },
        { index: 4,  value: d.estado_civil },
        { index: 5,  value: '' },
        { index: 6,  value: d.profissao },
        { index: 7,  value: '' },
        { index: 8,  value: d.rg },
        { index: 9,  value: d.cpf },
        { index: 10, value: d.rua },
        { index: 11, value: d.numero },
        { index: 12, value: '' },
        { index: 13, value: d.bairro },
        { index: 14, value: '' },
        { index: 15, value: d.cidade + ' ' },
        { index: 16, value: '- ' },
        { index: 17, value: d.estado },
        { index: 18, value: ', CEP' },
        { index: 19, value: ' ' },
        { index: 20, value: d.cep },
        { index: 21, value: d.email },
        { index: 22, value: dia },
        { index: 23, value: mes },
        { index: 24, value: nome },
        { index: 25, value: '' },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else {
      // CONTRATO — mapeamento exato do novo template
      const nome = d.nome.toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const orgao = d.orgao_expeditor || 'SSP/' + d.estado;
      const parcelas = fin.parcelas || [];
      const entrada = parcelas[0]?.valor || '';

      const mapa = [
        // Para 6 — qualificação cliente
        // RED[0]: NOME CLIENTE
        { index: 0,  value: nome },
        // RED[1-3]: brasileiro(a)
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: '' },
        { index: 3,  value: '' },
        // RED[4]: estado civil
        { index: 4,  value: d.estado_civil },
        // RED[5]: profissão
        { index: 5,  value: d.profissao },
        // RED[6]: RG número
        { index: 6,  value: d.rg },
        // RED[7]: órgão expedidor, RED[8]: espaço
        { index: 7,  value: orgao },
        { index: 8,  value: ' ' },
        // RED[9]: CPF
        { index: 9,  value: d.cpf },
        // RED[10]: endereço completo
        { index: 10, value: enderecoCompleto },
        // RED[11]: email
        { index: 11, value: d.email },

        // Para 13 — número do processo
        { index: 12, value: fin.numeroProcesso },

        // Para 19 — honorários
        // RED[13-17]: valor total (R$ 10.000,00)
        { index: 13, value: fin.valorTotal },
        { index: 14, value: '' },
        { index: 15, value: '' },
        { index: 16, value: '' },
        { index: 17, value: '' },
        // RED[18-23]: (dez mil reais)
        { index: 18, value: '(' + fin.valorTotalExtenso + ')' },
        { index: 19, value: '' },
        { index: 20, value: '' },
        { index: 21, value: '' },
        { index: 22, value: '' },
        { index: 23, value: '' },
        // RED[24-28]: valor entrada (5.000,00)
        { index: 24, value: entrada },
        { index: 25, value: '' },
        { index: 26, value: '' },
        { index: 27, value: '' },
        { index: 28, value: '' },
        // RED[29-32]: (cinco mil reais)
        { index: 29, value: '' },
        { index: 30, value: '' },
        { index: 31, value: '' },
        { index: 32, value: '' },
        // RED[33]: número de parcelas restantes
        { index: 33, value: fin.numParcelasRestantes },
        // RED[34-36]: (cinco) parcelas de R$
        { index: 34, value: '' },
        { index: 35, value: '' },
        { index: 36, value: '' },
        // RED[37-38]: valor da parcela (1.000,00)
        { index: 37, value: fin.valorParcela },
        { index: 38, value: '' },
        // RED[39-43]: (mil reais)
        { index: 39, value: '' },
        { index: 40, value: '' },
        { index: 41, value: '' },
        { index: 42, value: '' },
        { index: 43, value: '' },

        // Para 86 — data completa
        // RED[44]: 'São Paulo', RED[45]: ', '
        { index: 44, value: 'São Paulo' },
        { index: 45, value: ', ' },
        // RED[46-47]: dia (07)
        { index: 46, value: dia },
        { index: 47, value: '' },
        // RED[48-49]: ' de '
        { index: 48, value: ' ' },
        { index: 49, value: 'de ' },
        // RED[50]: mês
        { index: 50, value: mes + ' ' },
        // RED[51-54]: 'de 2026.'
        { index: 51, value: 'de ' },
        { index: 52, value: ano.substring(0, 3) },
        { index: 53, value: ano.substring(3) },
        { index: 54, value: '.' },

        // Para 91 — assinatura cliente
        { index: 55, value: nome },
      ];

      xml = substituirPorIndice(xml, mapa);
    }

    zip.file('word/document.xml', xml);
    const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    const nomeArq = tipo === 'procuracao'
      ? 'Procuracao_' + d.nome.replace(/\s+/g,'_') + '.docx'
      : 'Contrato_' + d.nome.replace(/\s+/g,'_') + '.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="' + nomeArq + '"');
    return res.status(200).send(buf);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
