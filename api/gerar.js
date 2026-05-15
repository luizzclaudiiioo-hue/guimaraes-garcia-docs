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

function removeRedColor(rPrXml) {
  // Remove the color element entirely to revert to default (black)
  return rPrXml.replace(/<w:color[^/]*\/>/g, '').replace(/<w:color[^>]*>[\s\S]*?<\/w:color>/g, '');
}

function substituir(xml, campos) {
  let idx = 0;
  return xml.replace(/(<w:r)([ >])([\s\S]*?)(<\/w:r>)/g, (match, tag, sep, inner, close) => {
    const rPrMatch = inner.match(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/);
    if (rPrMatch && isRed(rPrMatch[0])) {
      const tMatch = inner.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
      if (tMatch && tMatch[1].trim() && idx < campos.length) {
        const valor = (campos[idx] || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        idx++;
        // Remove red color from rPr and replace text
        const newRPr = rPrMatch[1] + removeRedColor(rPrMatch[2]) + rPrMatch[3];
        let newInner = inner.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, newRPr);
        newInner = newInner.replace(/<w:t[^>]*>[\s\S]*?<\/w:t>/, `<w:t xml:space="preserve">${valor}</w:t>`);
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
      const campos = [
        d.nome.toUpperCase(),        // NOME CLIENTE
        'brasileiro(a)',             // brasileir + o(a)
        d.estado_civil,              // solteiro(a)
        d.profissao,                 // profissão
        d.rg,                        // RG número
        d.orgao_expeditor || '',     // órgão expedidor
        d.cpf,                       // CPF
        d.rua,                       // Rua
        d.numero,                    // número
        d.bairro,                    // bairro
        d.cidade,                    // cidade
        d.estado,                    // estado
        d.cep,                       // CEP
        d.email,                     // email
        dia,                         // dia
        mes,                         // mês
        d.nome.toUpperCase(),        // NOME CLIENTE (assinatura)
      ];
      xml = substituir(xml, campos);
    } else {
      const parcelas = fin.parcelas || [];
      const metade = Math.ceil(parcelas.length / 2);
      const esq = parcelas.slice(0, metade);
      const dir = parcelas.slice(metade);
      const tabelaCampos = [];
      esq.forEach(p => tabelaCampos.push(p.label + ':', p.valor, p.data));
      dir.forEach(p => tabelaCampos.push(p.label + ':', p.valor, p.data));
      tabelaCampos.push('Total:', fin.valorTotal, '');

      const campos = [
        d.nome.toUpperCase(),
        'brasileiro(a)',
        d.estado_civil,
        d.profissao,
        d.cpf,
        d.rua + ', ' + d.numero + ', ' + d.bairro + ', ' + d.cidade + ' - ' + d.estado + ', CEP ' + d.cep,
        d.email,
        fin.numeroProcesso,
        fin.valorTotal + ' (' + fin.valorTotalExtenso + ')',
        parcelas[0]?.valor || '',
        fin.numParcelasRestantes + ' parcelas de ' + fin.valorParcela,
        dia,
        mes,
        d.nome.toUpperCase(),
        ...tabelaCampos,
      ];
      xml = substituir(xml, campos);
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
