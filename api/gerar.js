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

// Insere texto após um run específico no XML
function inserirTextoAposRun(xml, textoReferencia, novoTexto, formatoRun) {
  const escaped = textoReferencia.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`(<w:r[ >][\\s\\S]*?<w:t[^>]*>${escaped}<\\/w:t><\\/w:r>)`);
  return xml.replace(regex, `$1${formatoRun(novoTexto)}`);
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
      const parcelas = fin.parcelas || [];
      const nome = d.nome.toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const rgCompleto = d.rg + (d.orgao_expeditor ? ', expedido pela ' + d.orgao_expeditor : '');

      // Para 6: RED[0]=nome, RED[1-3]=brasileiro(a), RED[4]=estado_civil,
      // RED[5]=profissao, RED[6]=CPF, RED[7]=endereço, RED[8]=email
      // Inserimos RG no campo CPF: "portador(a) do RG nº XXXX, inscrito(a) no CPF sob o nº CPF"
      // O template tem ", inscrito(a) no CPF sob o nº" como texto fixo
      // então colocamos RG antes do CPF modificando o texto fixo via substituição direta

      // Substituir o texto fixo " inscrito(a) no CPF sob o nº" para incluir RG
      xml = xml.replace(
        /\xa0inscrito\(a\) no CPF sob o n[^<]*/g,
        `, portador(a) do RG nº ${esc(rgCompleto)}, inscrito(a) no CPF sob o nº`
      );

      // Honorários — RED[10-11]: valor total, RED[12-20]: extenso
      // RED[21-25]: entrada, RED[26-31]: entrada extenso
      // RED[30]: num parcelas, RED[31-33]: extenso
      // RED[34-35]: valor parcela, RED[36-43]: extenso
      // Para 91 — RED[41,42]: dia, RED[43,44]: mês
      // Para 96 — RED[45]: nome assinatura
      // Tabela — RED[46+]: parcelas

      const entrada = parcelas[0]?.valor || '';
      const numParc = fin.numParcelasRestantes;
      const valorParc = fin.valorParcela;

      const mapa = [
        // Para 6 — cliente
        { index: 0, value: nome },
        { index: 1, value: 'brasileiro(a)' },
        { index: 2, value: '' },
        { index: 3, value: '' },
        { index: 4, value: d.estado_civil },
        { index: 5, value: d.profissao },
        { index: 6, value: d.cpf },
        { index: 7, value: enderecoCompleto },
        { index: 8, value: d.email },
        // Para 13 — processo
        { index: 9, value: fin.numeroProcesso },
        // Para 19 — valor total (runs 10-11 = R$ VALOR, runs 12-20 = (extenso))
        { index: 10, value: fin.valorTotal },
        { index: 11, value: '' },
        { index: 12, value: '' },
        { index: 13, value: '' },
        { index: 14, value: '' },
        { index: 15, value: '(' + fin.valorTotalExtenso + ')' },
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
        // Para 91 — data: RED[41]='0', RED[42]='7' → dia / RED[43]=' ', RED[44]='abril '
        { index: 41, value: dia },
        { index: 42, value: '' },
        { index: 43, value: ' ' },
        { index: 44, value: mes + ' ' },
        // Para 96 — assinatura
        { index: 45, value: nome },
      ];

      xml = substituirPorIndice(xml, mapa);

      // Tabela — substituição direta por posição
      // Estrutura: 5 linhas, col 0-2 esquerda, col 4-6 direita
      // Linha 0: parcela[0] | parcela[metade]
      // Linha 1: parcela[1] | parcela[metade+1]
      // ...
      // Última linha direita col4='Total:', col5=valorTotal
      const metade = Math.ceil(parcelas.length / 2);
      const esq = parcelas.slice(0, metade);
      const dir = parcelas.slice(metade);

      let tIdx = 46;
      const tabelaMapa = [];

      for (let i = 0; i < esq.length; i++) {
        tabelaMapa.push({ index: tIdx++, value: esq[i].label + ':' });
        tabelaMapa.push({ index: tIdx++, value: esq[i].valor });
        tabelaMapa.push({ index: tIdx++, value: esq[i].data });
        if (dir[i]) {
          tabelaMapa.push({ index: tIdx++, value: dir[i].label + ':' });
          tabelaMapa.push({ index: tIdx++, value: dir[i].valor });
          tabelaMapa.push({ index: tIdx++, value: dir[i].data });
        } else {
          tabelaMapa.push({ index: tIdx++, value: 'Total:' });
          tabelaMapa.push({ index: tIdx++, value: fin.valorTotal });
          tIdx++;
        }
      }

      if (parcelas.length % 2 === 0) {
        tIdx += 3;
        tabelaMapa.push({ index: tIdx++, value: 'Total:' });
        tabelaMapa.push({ index: tIdx++, value: fin.valorTotal });
      }

      xml = substituirPorIndice(xml, tabelaMapa);
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
