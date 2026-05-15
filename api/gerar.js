import JSZip from 'jszip';

const GITHUB_USER = 'luizzclaudiiioo-hue';
const GITHUB_REPO = 'guimaraes-garcia-docs';
const GITHUB_BRANCH = 'main';

const EXTENSO_NUM = {
  1:'uma',2:'duas',3:'três',4:'quatro',5:'cinco',6:'seis',7:'sete',8:'oito',9:'nove',10:'dez',
  11:'onze',12:'doze',13:'treze',14:'quatorze',15:'quinze',16:'dezesseis',17:'dezessete',18:'dezoito',19:'dezenove',20:'vinte'
};

function numExtenso(n) { return EXTENSO_NUM[parseInt(n)] || n; }

function valorExtenso(valor) {
  if (!valor) return '';
  const num = parseFloat((valor+'').replace(/R\$\s*/,'').replace(/\./g,'').replace(',','.'));
  if (isNaN(num)) return valor;
  const centavos = Math.round((num % 1) * 100);
  const inteiro = Math.floor(num);
  const unidades = ['','um','dois','três','quatro','cinco','seis','sete','oito','nove',
    'dez','onze','doze','treze','quatorze','quinze','dezesseis','dezessete','dezoito','dezenove'];
  const dezenas = ['','','vinte','trinta','quarenta','cinquenta','sessenta','setenta','oitenta','noventa'];
  const centenas = ['','cem','duzentos','trezentos','quatrocentos','quinhentos',
    'seiscentos','setecentos','oitocentos','novecentos'];
  function menosDeMil(n) {
    if (n === 100) return 'cem';
    const c = Math.floor(n/100), r = n%100;
    if (r===0) return centenas[c];
    if (r<20) return (c>0?centenas[c]+' e ':'')+unidades[r];
    const d=Math.floor(r/10),u=r%10;
    return (c>0?centenas[c]+' e ':'')+dezenas[d]+(u>0?' e '+unidades[u]:'');
  }
  function porExtenso(n) {
    if (n===0) return 'zero';
    if (n<1000) return menosDeMil(n);
    if (n<1000000) {
      const mil=Math.floor(n/1000), r=n%1000;
      const milStr = mil===1?'mil':menosDeMil(mil)+' mil';
      return r===0?milStr:milStr+' e '+menosDeMil(r);
    }
    return String(n);
  }
  const parteInteira = porExtenso(inteiro);
  const moeda = inteiro===1?'real':'reais';
  if (centavos===0) return parteInteira+' '+moeda;
  return parteInteira+' '+moeda+' e '+menosDeMil(centavos)+(centavos===1?' centavo':' centavos');
}

async function fetchTemplate(filename) {
  const url = `https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/${GITHUB_BRANCH}/templates/${filename}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Template not found: ${filename} (${res.status})`);
  return Buffer.from(await res.arrayBuffer());
}

function isRed(rPrXml) { return /<w:color[^>]*w:val="FF0000"/.test(rPrXml); }

function removeRedColor(rPrInner) {
  return rPrInner.replace(/<w:color[^/]*\/>/g,'').replace(/<w:color[^>]*>[\s\S]*?<\/w:color>/g,'');
}

function esc(s) { return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

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

    const filenameMap = {
      procuracao: 'template_procuracao.docx',
      contrato: 'template_contrato.docx',
      declaracao: 'template_declaracao.docx',
      revogacao: 'template_revogacao.docx',
      residencia_propria: 'template_residencia_propria.docx',
      amaisa: 'template_amaisa.docx',
      proprietario: 'template_proprietario.docx',
    };
    const filename = filenameMap[tipo];
    if (!filename) return res.status(400).json({ error: `Tipo desconhecido: ${tipo}` });
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
        // RED[10]: endereço completo (template unificado)
        { index: 10, value: d.rua + ', ' + d.numero + ', ' + d.bairro + ', ' + d.cidade + ' - ' + d.estado + ', CEP ' + d.cep },
        // RED[11]: email (antes era RED[21])
        { index: 11, value: d.email },
        // RED[12]: dia
        { index: 12, value: dia },
        // RED[13]: mês
        { index: 13, value: mes },
        // RED[14]: nome assinatura
        { index: 14, value: nome },
        { index: 15, value: '' },
      ];
      xml = substituirPorIndice(xml, mapa);

      // Remove vírgulas duplicadas geradas por campos vazios no endereço
      xml = xml.replace(/,\s*,\s*,/g, ',').replace(/,\s*,/g, ',');

    } else {
      const nome = d.nome.toUpperCase();
      const orgao = d.orgao_expeditor || ('SSP/' + d.estado);
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const entrada = fin.valorEntrada || '';
      const numParc = fin.numParcelasRestantes;
      const numParcExtenso = numExtenso(numParc);
      const stripRS = v => (v||'').replace(/^R\$\s*/,'');

      // Extensos automáticos
      const entradaExtenso = valorExtenso(entrada);
      const parcelaExtenso = valorExtenso(fin.valorParcela);

      const mapa = [
        // Para 6
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: '' },
        { index: 3,  value: '' },
        { index: 4,  value: d.estado_civil },
        { index: 5,  value: d.profissao },
        { index: 6,  value: d.rg },
        { index: 7,  value: orgao + ',' },
        { index: 8,  value: ' ' },
        { index: 9,  value: d.cpf },
        { index: 10, value: enderecoCompleto },
        { index: 11, value: d.email },
        // Para 13
        { index: 12, value: fin.numeroProcesso },
        // Para 19 — valor total
        // RED[13]='R$' RED[14]=' ' RED[15-17]=número
        { index: 13, value: 'R$' },
        { index: 14, value: ' ' },
        { index: 15, value: stripRS(fin.valorTotal) },
        { index: 16, value: '' },
        { index: 17, value: '' },
        // RED[18]=' (' RED[19-23]=extenso
        { index: 18, value: ' (' },
        { index: 19, value: valorExtenso(fin.valorTotal) + ')' },
        { index: 20, value: '' },
        { index: 21, value: '' },
        { index: 22, value: '' },
        { index: 23, value: '' },
        // RED[24-28]=valor entrada
        { index: 24, value: stripRS(entrada) },
        { index: 25, value: '' },
        { index: 26, value: '' },
        { index: 27, value: '' },
        { index: 28, value: ' (' },
        // RED[29-32]=extenso entrada — automático
        { index: 29, value: entradaExtenso + ')' },
        { index: 30, value: '' },
        { index: 31, value: '' },
        { index: 32, value: '' },
        // RED[33]=num parcelas
        { index: 33, value: numParc },
        // RED[34]=' ('
        { index: 34, value: ' (' },
        // RED[35]=extenso num parcelas — automático
        { index: 35, value: numParcExtenso },
        // RED[36]=') parcelas de R$ '
        { index: 36, value: ') parcelas de R$ ' },
        // RED[37-38]=valor parcela
        { index: 37, value: stripRS(fin.valorParcela) },
        { index: 38, value: '' },
        // RED[39-43]=extenso parcela — automático
        { index: 39, value: ' (' },
        { index: 40, value: parcelaExtenso + ')' },
        { index: 41, value: '' },
        { index: 42, value: '' },
        { index: 43, value: '' },
        // Para 86 — data
        { index: 44, value: 'São Paulo' },
        { index: 45, value: ', ' },
        { index: 46, value: dia },
        { index: 47, value: '' },
        { index: 48, value: ' ' },
        { index: 49, value: 'de ' },
        { index: 50, value: mes + ' ' },
        { index: 51, value: 'de ' },
        { index: 52, value: ano.substring(0,3) },
        { index: 53, value: ano.substring(3) },
        { index: 54, value: '.' },
        // Para 91 — assinatura
        { index: 55, value: nome },
      ];

      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'declaracao') {
      const nome = d.nome.toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: d.estado_civil },
        { index: 3,  value: d.profissao },
        { index: 4,  value: d.rg },
        { index: 5,  value: d.cpf },
        { index: 6,  value: enderecoCompleto },
        { index: 7,  value: d.email },
        { index: 8,  value: 'São Paulo, ' },
        { index: 9,  value: dia },
        { index: 10, value: ' de ' },
        { index: 11, value: mes + ' ' },
        { index: 12, value: 'de ' + ano.substring(0,3) },
        { index: 13, value: ano.substring(3) },
        { index: 14, value: '.' },
        { index: 15, value: nome },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'revogacao') {
      const nome = d.nome.toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: d.estado_civil },
        { index: 3,  value: d.profissao },
        { index: 4,  value: d.rg },
        { index: 5,  value: d.cpf },
        { index: 6,  value: enderecoCompleto },
        { index: 7,  value: d.email },
        { index: 8,  value: fin.numeroProcesso || '' },
        { index: 9,  value: 'São Paulo, ' + dia + ' de ' + mes + ' de ' + ano + '.' },
        { index: 10, value: nome },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'residencia_propria') {
      const nome = d.nome.toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: d.estado_civil },
        { index: 3,  value: d.profissao },
        { index: 4,  value: d.rg },
        { index: 5,  value: 'SSP/' + d.estado },
        { index: 6,  value: d.cpf },
        { index: 7,  value: enderecoCompleto },
        { index: 8,  value: d.email },
        { index: 9,  value: enderecoCompleto },
        { index: 10, value: '.' },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'amaisa') {
      const nome = d.nome.toUpperCase();
      const nomeAmaisa = (fin.nomeAmaisa || '').toUpperCase();
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      // Calcular tempo de união a partir da data
      let anosUniao = '', mesesUniao = '';
      if (fin.dataUniao) {
        const [diaU, mesU, anoU] = fin.dataUniao.split('/').map(Number);
        const inicio = new Date(anoU, mesU - 1, diaU);
        const agora = new Date();
        let diffMeses = (agora.getFullYear() - inicio.getFullYear()) * 12 + (agora.getMonth() - inicio.getMonth());
        anosUniao = String(Math.floor(diffMeses / 12));
        mesesUniao = String(diffMeses % 12);
      }
      const mapa = [
        // Dados do(a) companheiro(a) - RED[0-6]
        { index: 0,  value: nomeAmaisa.split(' ')[0] + ' ' },
        { index: 1,  value: nomeAmaisa.split(' ').slice(1).join(' ') },
        { index: 2,  value: d.rg },
        { index: 3,  value: 'SSP/' + d.estado },
        { index: 4,  value: d.cpf },
        { index: 5,  value: enderecoCompleto },
        { index: 6,  value: d.email },
        // Tempo de união - RED[7-13]
        { index: 7,  value: anosUniao },
        { index: 8,  value: ' ' },
        { index: 9,  value: anosUniao === '1' ? 'ano' : 'anos' },
        { index: 10, value: ' ' },
        { index: 11, value: 'e ' },
        { index: 12, value: mesesUniao },
        { index: 13, value: ' meses' },
        // Dados do cliente - RED[14-23]
        { index: 14, value: nome.split(' ')[0] },
        { index: 15, value: ' ' + nome.split(' ').slice(1).join(' ') },
        { index: 16, value: d.rg },
        { index: 17, value: 'SSP' },
        { index: 18, value: '/' },
        { index: 19, value: d.estado },
        { index: 20, value: d.cpf },
        { index: 21, value: '' },
        { index: 22, value: '' },
        { index: 23, value: '' },
        // Data - RED[24-33]
        { index: 24, value: 'São Paulo, ' },
        { index: 25, value: dia },
        { index: 26, value: ' de' },
        { index: 27, value: ' ' },
        { index: 28, value: mes },
        { index: 29, value: ' ' },
        { index: 30, value: 'de' },
        { index: 31, value: ' ' },
        { index: 32, value: ano.substring(0,3) },
        { index: 33, value: ano.substring(3) },
        { index: 34, value: '' },
        // Assinatura - RED[35-36]
        { index: 35, value: nome.split(' ')[0] },
        { index: 36, value: ' ' + nome.split(' ').slice(1).join(' ') },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'proprietario') {
      const nomeLocador = d.nome.toUpperCase();
      const nomeLocatario = (fin.nomeLocatario || '').toUpperCase();
      const cpfLocatario = fin.cpfLocatario || '';
      const enderecoLocador = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const orgao = d.orgao_expeditor || ('SSP/' + d.estado);
      const mapa = [
        { index: 0,  value: nomeLocador.split(' ')[0] + ' ' },
        { index: 1,  value: nomeLocador.split(' ').slice(1).join(' ') },
        { index: 2,  value: 'brasileiro(a)' },
        { index: 3,  value: d.rg },
        { index: 4,  value: orgao },
        { index: 5,  value: d.cpf },
        { index: 6,  value: enderecoLocador },
        // Endereço do imóvel - RED[7-9] (fixo no template, mantém texto original)
        { index: 7,  value: 'Rua Barra Funda, ' },
        { index: 8,  value: '' },
        { index: 9,  value: ', Barra Funda, São Paulo – SP, CEP 000' },
        // Locatário - RED[10-11]
        { index: 10, value: nomeLocatario },
        { index: 11, value: cpfLocatario },
        // Data - RED[12-18]
        { index: 12, value: 'São Paulo, ' },
        { index: 13, value: dia },
        { index: 14, value: ' de ' },
        { index: 15, value: mes },
        { index: 16, value: ' de' },
        { index: 17, value: ' ' + ano },
        { index: 18, value: '.' },
        // Assinatura - RED[19]
        { index: 19, value: nomeLocador },
      ];
      xml = substituirPorIndice(xml, mapa);
    }

    zip.file('word/document.xml', xml);
    const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    const nomesArquivo = {
      procuracao: 'Procuracao_',
      contrato: 'Contrato_',
      declaracao: 'Declaracao_Hipossuficiencia_',
      revogacao: 'Revogacao_Mandato_',
      residencia_propria: 'Declaracao_Residencia_',
      amaisa: 'Declaracao_Amaisa_',
      proprietario: 'Declaracao_Proprietario_',
    };
    const nomeArq = (nomesArquivo[tipo] || 'Documento_') + d.nome.replace(/\s+/g,'_') + '.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="' + nomeArq + '"');
    return res.status(200).send(buf);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
