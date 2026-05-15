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
  return rPrInner.replace(/<w:color[^/]*\/>/g, '').replace(/<w:color[^>]*>[\s\S]*?<\/w:color>/g, '');
}

// Replace red runs using a specific mapping array
// Each entry: { index: N, value: 'text' } — replaces the Nth red run (0-based)
function substituirPorIndice(xml, mapa) {
  let redIdx = 0;
  return xml.replace(/(<w:r)([ >])([\s\S]*?)(<\/w:r>)/g, (match, tag, sep, inner, close) => {
    const rPrMatch = inner.match(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/);
    if (rPrMatch && isRed(rPrMatch[0])) {
      const currentIdx = redIdx++;
      const entry = mapa.find(m => m.index === currentIdx);
      if (entry !== undefined) {
        const valor = (entry.value || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        const newRPr = rPrMatch[1] + removeRedColor(rPrMatch[2]) + rPrMatch[3];
        let newInner = inner.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, newRPr);
        newInner = newInner.replace(/<w:t[^>]*>[\s\S]*?<\/w:t>/, `<w:t xml:space="preserve">${valor}</w:t>`);
        return tag + sep + newInner + close;
      } else {
        // Just remove red color but keep text
        const newRPr = rPrMatch[1] + removeRedColor(rPrMatch[2]) + rPrMatch[3];
        const newInner = inner.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, newRPr);
        return tag + sep + newInner + close;
      }
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
      // Mapeamento exato dos runs vermelhos (índice baseado na ordem no documento):
      // 0: NOME, 1: ' CLIENTE', 2: 'brasileir', 3: 'o(a)', 4: 'solteiro', 5: '(a)',
      // 6: 'profissão', 7: ' XXX', 8: '0000'(RG), 9: '0000'(CPF),
      // 10: 'Rua', 11: 'XXXX'(numero), 12: '0000'(numero), 13: 'bairro', 14: ' XXX',
      // 15: 'São Paulo ', 16: '- ', 17: 'SP', 18: ', CEP', 19: ' ', 20: '0000'(CEP),
      // 21: 'XXX'(email), 22: 'dia', 23: 'mês', 24: 'NOME', 25: ' CLIENTE'

      const nome = d.nome.toUpperCase();
      const mapa = [
        { index: 0, value: nome },          // NOME
        { index: 1, value: '' },             // ' CLIENTE' — já no run anterior
        { index: 2, value: 'brasileiro(a)' },// brasileir + o(a) merged
        { index: 3, value: '' },             // o(a) — já no anterior
        { index: 4, value: d.estado_civil }, // solteiro
        { index: 5, value: '' },             // (a) — já no anterior
        { index: 6, value: d.profissao },    // profissão
        { index: 7, value: '' },             // XXX — já no anterior
        { index: 8, value: d.rg + (d.orgao_expeditor ? ', expedido pela ' + d.orgao_expeditor : '') }, // RG
        { index: 9, value: d.cpf },          // CPF
        { index: 10, value: d.rua },         // Rua
        { index: 11, value: d.numero },      // número
        { index: 12, value: '' },            // numero duplicado
        { index: 13, value: d.bairro },      // bairro
        { index: 14, value: '' },            // XXX bairro complemento
        { index: 15, value: d.cidade + ' ' },// cidade
        { index: 16, value: '- ' },          // separador
        { index: 17, value: d.estado },      // estado
        { index: 18, value: ', CEP' },       // , CEP
        { index: 19, value: ' ' },           // espaço
        { index: 20, value: d.cep },         // CEP número
        { index: 21, value: d.email },       // email
        { index: 22, value: dia },           // dia
        { index: 23, value: mes },           // mês
        { index: 24, value: nome },          // NOME assinatura
        { index: 25, value: '' },            // CLIENTE assinatura
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

      // Sequential replacement for contract
      let idx = 0;
      const camposContrato = [
        d.nome.toUpperCase(), '', 'brasileiro(a)', '', d.estado_civil, '', d.profissao, '',
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
