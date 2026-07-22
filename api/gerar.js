const JSZip = require('jszip');

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

function isRed(rPrXml) { return /<w:color[^>]*w:val="(FF0000|EE0000)"/.test(rPrXml); }

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

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { tipo, dados: d, financeiro: fin } = req.body;
    const agora = new Date();
    const spOffset = -3 * 60;
    const spTime = new Date(agora.getTime() + (spOffset - agora.getTimezoneOffset()) * 60000);
    const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
    const dia = String(spTime.getDate());
    const mes = meses[spTime.getMonth()];
    const ano = String(spTime.getFullYear());

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
      // RED[0]=nome, [1]=nacionalidade, [2]=estado_civil, [3]=profissao,
      // [4]=RG, [5]=orgao, [6]=CPF, [7]=endereco, [8]=email,
      // [9]=dia, [10]=mes, [11]=ano, [12]=nome assinatura
      const nome = d.nome.toUpperCase();
      const orgao = d.orgao_expeditor || ('SSP/' + d.estado);
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro' },
        { index: 2,  value: d.estado_civil ? d.estado_civil.replace('(a)','').trim() : 'solteiro' },
        { index: 3,  value: d.profissao },
        { index: 4,  value: d.rg },
        { index: 5,  value: orgao },
        { index: 6,  value: d.cpf },
        { index: 7,  value: enderecoCompleto },
        { index: 8,  value: d.email },
        { index: 9,  value: dia },
        { index: 10, value: mes },
        { index: 11, value: ano },
        { index: 12, value: nome },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'contrato') {
      // RED[0]=nome, [1]=nacionalidade, [2]=estado_civil, [3]=profissao,
      // [4]=RG, [5]=orgao, [6]=CPF, [7]=endereco, [8]=email,
      // [9]=numeroProcesso, [10]=valorTotal, [11]=valorEntrada,
      // [12]=numParcelas, [13]=valorParcela,
      // [14]=dia, [15]=mes, [16]=ano, [17]=nome assinatura
      const nome = d.nome.toUpperCase();
      const orgao = d.orgao_expeditor || ('SSP/' + d.estado);
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const numParc = fin.numParcelasRestantes || '';
      const numParcExtenso = numExtenso(numParc);
      const stripRS = v => (v||'').replace(/^R\$\s*/,'');
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro' },
        { index: 2,  value: d.estado_civil ? d.estado_civil.replace('(a)','').trim() : 'solteiro' },
        { index: 3,  value: d.profissao },
        { index: 4,  value: d.rg },
        { index: 5,  value: orgao },
        { index: 6,  value: d.cpf },
        { index: 7,  value: enderecoCompleto },
        { index: 8,  value: d.email },
        { index: 9,  value: fin.numeroProcesso || '' },
        { index: 10, value: `R$ ${stripRS(fin.valorTotal)} (${valorExtenso(fin.valorTotal)})` },
        { index: 11, value: `R$ ${stripRS(fin.valorEntrada)} (${valorExtenso(fin.valorEntrada)})` },
        { index: 12, value: `${numParc} (${numParcExtenso})` },
        { index: 13, value: `R$ ${stripRS(fin.valorParcela)} (${valorExtenso(fin.valorParcela)})` },
        { index: 14, value: dia },
        { index: 15, value: mes },
        { index: 16, value: ano },
        { index: 17, value: nome },
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
      const orgao = d.orgao_expeditor || ('SSP/' + d.estado);
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const mapa = [
        { index: 0,  value: nome },
        { index: 1,  value: 'brasileiro(a)' },
        { index: 2,  value: d.estado_civil },
        { index: 3,  value: d.profissao },
        { index: 4,  value: d.rg },
        { index: 5,  value: orgao },
        { index: 6,  value: d.cpf },
        { index: 7,  value: enderecoCompleto },
        { index: 8,  value: d.email },
        { index: 9,  value: enderecoCompleto },
        { index: 10, value: '.' },
        { index: 11, value: `${d.cidade || 'São Paulo'}, ${dia} de ${mes} de ${ano}.` },
        { index: 12, value: nome },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'amaisa') {
      const nomeEsposa = d.nome.toUpperCase();
      const nomePreso = (fin.nomePreso || '').toUpperCase();
      const orgao = d.orgao_expeditor || ('SSP/' + d.estado);
      const enderecoCompleto = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      let anosUniao = '0', mesesUniao = '0';
      if (fin.dataUniao) {
        const partes = fin.dataUniao.split('/');
        if (partes.length === 3) {
          const inicio = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
          const agora2 = new Date();
          let diffMeses = (agora2.getFullYear() - inicio.getFullYear()) * 12 + (agora2.getMonth() - inicio.getMonth());
          anosUniao = String(Math.floor(diffMeses / 12));
          mesesUniao = String(diffMeses % 12);
        }
      }
      const orgaoPreso = fin.orgaoPreso || 'SSP/SP';
      const orgaoPartes = orgaoPreso.split('/');
      const orgaoParte1 = orgaoPartes[0] || 'SSP';
      const orgaoParte2 = orgaoPartes[1] || 'SP';
      const mapa = [
        { index: 0,  value: nomeEsposa.split(' ')[0] + ' ' },
        { index: 1,  value: nomeEsposa.split(' ').slice(1).join(' ') },
        { index: 2,  value: d.rg },
        { index: 3,  value: orgao },
        { index: 4,  value: d.cpf },
        { index: 5,  value: enderecoCompleto },
        { index: 6,  value: d.email },
        { index: 7,  value: anosUniao },
        { index: 8,  value: ' ' },
        { index: 9,  value: Number(anosUniao) === 1 ? 'ano' : 'anos' },
        { index: 10, value: ' ' },
        { index: 11, value: 'e ' },
        { index: 12, value: mesesUniao },
        { index: 13, value: ' meses' },
        { index: 14, value: nomePreso.split(' ')[0] },
        { index: 15, value: ' ' + nomePreso.split(' ').slice(1).join(' ') },
        { index: 16, value: '' },
        { index: 17, value: fin.rgPreso || '' },
        { index: 18, value: orgaoParte1 },
        { index: 19, value: '/' },
        { index: 20, value: orgaoParte2 },
        { index: 21, value: fin.cpfPreso || '' },
        { index: 22, value: fin.localDetencao ? fin.localDetencao.split(',')[0] || '' : '' },
        { index: 23, value: fin.localDetencao ? (fin.localDetencao.split(',')[1] || '').trim() : '' },
        { index: 24, value: fin.matricula || '' },
        { index: 25, value: 'São Paulo, ' },
        { index: 26, value: dia },
        { index: 27, value: ' de' },
        { index: 28, value: ' ' },
        { index: 29, value: mes },
        { index: 30, value: ' ' },
        { index: 31, value: 'de' },
        { index: 32, value: ' ' },
        { index: 33, value: ano.substring(0,3) },
        { index: 34, value: ano.substring(3) },
        { index: 35, value: '' },
        { index: 36, value: nomeEsposa.split(' ')[0] },
        { index: 37, value: ' ' + nomeEsposa.split(' ').slice(1).join(' ') },
      ];
      xml = substituirPorIndice(xml, mapa);

    } else if (tipo === 'proprietario') {
      const nomeLocador = d.nome.toUpperCase();
      const orgaoLocador = d.orgao_expeditor || ('SSP/' + d.estado);
      const enderecoLocador = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;
      const enderecoImovel = fin.enderecoImovel || '';
      const nomeLocatario = (fin.nomeLocatario || '').toUpperCase();
      const cpfLocatario = fin.cpfLocatario || '';
      const mapa = [
        { index: 0,  value: nomeLocador.split(' ')[0] + ' ' },
        { index: 1,  value: nomeLocador.split(' ').slice(1).join(' ') },
        { index: 2,  value: 'brasileiro(a)' },
        { index: 3,  value: d.rg },
        { index: 4,  value: orgaoLocador },
        { index: 5,  value: d.cpf },
        { index: 6,  value: enderecoLocador },
        { index: 7,  value: enderecoImovel },
        { index: 8,  value: '' },
        { index: 9,  value: '' },
        { index: 10, value: nomeLocatario },
        { index: 11, value: cpfLocatario },
        { index: 12, value: 'São Paulo, ' },
        { index: 13, value: dia },
        { index: 14, value: ' de ' },
        { index: 15, value: mes },
        { index: 16, value: ' de' },
        { index: 17, value: ' ' + ano },
        { index: 18, value: '.' },
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
