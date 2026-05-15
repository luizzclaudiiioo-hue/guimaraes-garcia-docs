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
      // Mapeamento exato baseado na análise do template:
      // RED[0]: 'NOME', RED[1]: ' CLIENTE'
      // RED[2]: 'brasileir', RED[3]: 'o(a)'
      // RED[4]: 'solteiro', RED[5]: '(a)'
      // RED[6]: 'profissão', RED[7]: ' XXX'
      // RED[8]: '0000' (RG número — template já tem ", expedido pela SSP/SP," fixo)
      // RED[9]: '0000' (CPF)
      // RED[10]: 'Rua' (nome da rua)
      // RED[11]: 'XXXX' (número do endereço)
      // RED[12]: '0000' (campo extra — esvaziar)
      // RED[13]: 'bairro', RED[14]: ' XXX' (bairro nos dois runs)
      // RED[15]: 'São Paulo ' (cidade), RED[16]: '- ', RED[17]: 'SP' (estado)
      // RED[18]: ', CEP', RED[19]: ' ', RED[20]: '0000' (CEP)
      // RED[21]: 'XXX' (email)
      // RED[22]: 'dia', RED[23]: 'mês' (data)
      // RED[24]: 'NOME', RED[25]: ' CLIENTE' (assinatura)

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
      const parcelas = fin.parcelas || [];
      const metade = Math.ceil(parcelas.length / 2);
      const esq = parcelas.slice(0, metade);
      const dir = parcelas.slice(metade);
      const tabelaCampos = [];
      esq.forEach(p => tabelaCampos.push(p.label + ':', p.valor, p.data));
      dir.forEach(p => tabelaCampos.push(p.label + ':', p.valor, p.data));
      tabelaCampos.push('Total:', fin.valorTotal, '');

      const camposContrato = [
        d.nome.toUpperCase(), '',
        'brasileiro(a)', '',
        d.estado_civil, '',
        d.profissao, '',
        d.cpf,
        d.rua + ', ' + d.numero + ', ' + d.bairro + ', ' + d.cidade + ' - ' + d.estado + ', CEP ' + d.cep,
        d.email,
        fin.numeroProcesso,
        fin.valorTotal + ' (' + fin.valorTotalExtenso + ')',
        parcelas[0]?.valor || '',
        fin.numParcelasRestantes + ' parcelas de ' + fin.valorParcela,
        dia, mes,
        d.nome.toUpperCase(), '',
        ...tabelaCampos,
      ];

      let idx = 0;
      xml = xml.replace(/(<w:r)([ >])([\s\S]*?)(<\/w:r>)/g, (match, tag, sep, inner, close) => {
        const rPrMatch = inner.match(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/);
        if (rPrMatch && isRed(rPrMatch[0])) {
          const valor = (camposContrato[idx++] || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
          const newRPr = rPrMatch[1] + removeRedColor(rPrMatch[2]) + rPrMatch[3];
          let newInner = inner.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, newRPr);
          newInner = newInner.replace(/<w:t[^>]*>[\s\S]*?<\/w:t>/, `<w:t xml:space="preserve">${valor}</w:t>`);
          return tag + sep + newInner + close;
        }
        return match;
      });
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
