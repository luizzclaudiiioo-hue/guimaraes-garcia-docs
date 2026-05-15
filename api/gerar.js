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

function substituirPorIndice(xml, mapa) {
  let redIdx = 0;
  return xml.replace(/(<w:r)([ >])([\s\S]*?)(<\/w:r>)/g, (match, tag, sep, inner, close) => {
    const rPrMatch = inner.match(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/);
    if (rPrMatch && isRed(rPrMatch[0])) {
      const currentIdx = redIdx++;
      const entry = mapa.find(m => m.index === currentIdx);
      const valor = entry !== undefined
        ? (entry.value || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
        : '';
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
      // Contrato — mapeamento exato baseado na análise do template:
      // Para 6 — dados cliente:
      // RED[0]: nome, RED[1]: 'brasileir', RED[2]: 'o', RED[3]: '(a)',
      // RED[4]: estado_civil, RED[5]: profissão, RED[6]: CPF,
      // RED[7]: endereço completo, RED[8]: email
      // Para 13 — processo: RED[9]: número processo
      // Para 19 — honorários (muitos runs):
      // RED[10-14]: valor total (ex: R$ 10.000,00)
      // RED[15-20]: valor total por extenso (ex: dez mil reais)
      // RED[21-25]: valor entrada
      // RED[26-29]: valor entrada por extenso
      // RED[30]: num parcelas
      // RED[31-33]: num parcelas por extenso
      // RED[34-35]: valor parcela
      // RED[36-40]: valor parcela por extenso
      // Para 91 — data: RED[41]: dia1, RED[42]: dia2, RED[43]: espaço, RED[44]: mês
      // Para 96 — assinatura: RED[45]: nome
      // Tabela: RED[46-65]: parcelas e total

      const parcelas = fin.parcelas || [];

      // Montar valor total como string com runs separados
      // O template tem o valor em múltiplos runs: 'R$', ' ', '10', '.0', '00,00'
      // Vamos colocar tudo no primeiro run e zerar os demais
      const valorTotalStr = fin.valorTotal; // ex: R$ 10.000,00
      const valorTotalExtensoStr = fin.valorTotalExtenso; // ex: dez mil reais
      const entrada = parcelas[0] ? parcelas[0].valor : '';
      const numParc = fin.numParcelasRestantes;
      const valorParc = fin.valorParcela;

      const nome = d.nome.toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;

      const mapa = [
        // Para 6 — cliente
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: '' },
        { index: 3,  value: '' },
        { index: 4,  value: d.estado_civil },
        { index: 5,  value: d.profissao },
        { index: 6,  value: d.cpf },
        { index: 7,  value: enderecoCompleto },
        { index: 8,  value: d.email },
        // Para 13 — processo
        { index: 9,  value: fin.numeroProcesso },
        // Para 19 — honorários (valor total em múltiplos runs)
        { index: 10, value: valorTotalStr },
        { index: 11, value: '' },
        { index: 12, value: '' },
        { index: 13, value: '' },
        { index: 14, value: '' },
        { index: 15, value: ' (' + valorTotalExtensoStr + ')' },
        { index: 16, value: '' },
        { index: 17, value: '' },
        { index: 18, value: '' },
        { index: 19, value: '' },
        { index: 20, value: '' },
        // entrada
        { index: 21, value: entrada },
        { index: 22, value: '' },
        { index: 23, value: '' },
        { index: 24, value: '' },
        { index: 25, value: '' },
        { index: 26, value: '' },
        { index: 27, value: '' },
        { index: 28, value: '' },
        { index: 29, value: '' },
        // num parcelas
        { index: 30, value: numParc },
        { index: 31, value: '' },
        { index: 32, value: '' },
        { index: 33, value: '' },
        // valor parcela
        { index: 34, value: valorParc },
        { index: 35, value: '' },
        { index: 36, value: '' },
        { index: 37, value: '' },
        { index: 38, value: '' },
        { index: 39, value: '' },
        { index: 40, value: '' },
        // Para 91 — data
        { index: 41, value: dia },
        { index: 42, value: '' },
        { index: 43, value: ' ' },
        { index: 44, value: mes + ' ' },
        // Para 96 — assinatura
        { index: 45, value: nome },
      ];

      xml = substituirPorIndice(xml, mapa);

      // Tabela — substituir diretamente no XML da tabela
      // RED[46-65] são os campos da tabela
      // Vamos fazer uma segunda passagem só para a tabela
      // Já foram zerados pelo substituirPorIndice acima (valor ''),
      // então precisamos preencher com os dados das parcelas
      
      // Reprocessar a tabela com dados corretos
      const metade = Math.ceil(parcelas.length / 2);
      const esq = parcelas.slice(0, metade);
      const dir = parcelas.slice(metade);

      // Build table replacement map starting at RED[46]
      const tabelaMapa = [];
      let tIdx = 46;
      
      // Row 0: esq[0] cols 0,1,2 | dir[0] cols 4,5,6
      // Row 1: esq[1] | dir[1]
      // etc.
      // Last row right: Total col5
      
      esq.forEach((p, i) => {
        tabelaMapa.push({ index: tIdx++, value: p.label + ':' });
        tabelaMapa.push({ index: tIdx++, value: p.valor });
        tabelaMapa.push({ index: tIdx++, value: p.data });
        if (dir[i]) {
          tabelaMapa.push({ index: tIdx++, value: dir[i].label + ':' });
          tabelaMapa.push({ index: tIdx++, value: dir[i].valor });
          tabelaMapa.push({ index: tIdx++, value: dir[i].data });
        } else {
          tabelaMapa.push({ index: tIdx++, value: 'Total:' });
          tabelaMapa.push({ index: tIdx++, value: fin.valorTotal });
          tIdx++; // empty
        }
      });

      // If even number of parcelas, add total row
      if (parcelas.length % 2 === 0) {
        tIdx += 3; // skip left cols
        tabelaMapa.push({ index: tIdx++, value: 'Total:' });
        tabelaMapa.push({ index: tIdx++, value: fin.valorTotal });
      }

      xml = substituirPorIndice(xml, tabelaMapa);
    }

    zip.file('word/document.xml', xml);
    const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    const nome = tipo === 'procuracao'
      ? 'Procuracao_' + d.nome.replace(/\s+/g,'_') + '.docx'
      : 'Contrato_' + d.nome.replace(/\s+/g,'_') + '.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="' + nome + '"');
    return res.status(200).send(buf);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
