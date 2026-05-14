import { useState } from "react";

const ANTHROPIC_MODEL = "claude-sonnet-4-5";


const EXTRACTION_PROMPT = `Você é um assistente jurídico. Extraia os dados do cliente a partir do texto abaixo, que foi enviado via WhatsApp em resposta a um script de coleta de dados para elaboração de documentos jurídicos.

Retorne APENAS um JSON válido, sem markdown, sem texto adicional, com exatamente estas chaves:
{
  "nome": "",
  "estado_civil": "",
  "profissao": "",
  "rg": "",
  "orgao_expeditor": "",
  "cpf": "",
  "rua": "",
  "numero": "",
  "bairro": "",
  "cidade": "",
  "estado": "",
  "cep": "",
  "email": ""
}

Se algum dado não for encontrado, deixe o campo com string vazia.
Para RG e órgão expedidor: separe o número do RG do órgão (ex: "12.345.678-9 SSP/SP" → rg: "12.345.678-9", orgao_expeditor: "SSP/SP").
Para estado civil e profissão: normalize para minúsculas.

Texto do cliente:`;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function dataHoje() {
  const d = new Date();
  const meses = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  return { dia: d.getDate(), mes: meses[d.getMonth()], ano: d.getFullYear() };
}

function parcelasVazias(n) {
  return Array.from({ length: n }, (_, i) => ({
    label: i === 0 ? "Entrada" : `${i}ª parcela`,
    valor: "",
    data: "",
  }));
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ─── Geração DOCX: Procuração ─────────────────────────────────────────────────

async function gerarProcuracao(d) {
  const { Document, Packer, Paragraph, TextRun, AlignmentType } = await import("https://cdn.skypack.dev/docx@8.5.0");
  const { dia, mes, ano } = dataHoje();
  const B = (t) => new TextRun({ text: t, bold: true, font: "Times New Roman", size: 24 });
  const N = (t) => new TextRun({ text: t, font: "Times New Roman", size: 24 });
  const rgCompleto = d.orgao_expeditor ? `${d.rg}, expedido pela ${d.orgao_expeditor}` : d.rg;
  const endereco = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 } } },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("GUIMARÃES & GARCIA")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B("SOCIEDADE DE ADVOGADOS")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B("PROCURAÇÃO")] }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 400 }, indent: { firstLine: 720 },
          children: [
            B(d.nome.toUpperCase()),
            N(`, ${d.estado_civil}, ${d.profissao}, portador(a) do RG nº ${rgCompleto}, inscrito(a) no CPF sob o nº ${d.cpf}, residente e domiciliado(a) na ${endereco}, com endereço eletrônico: ${d.email}, abaixo assinado, pelo presente instrumento de mandato, nomeia e constitui como seus procuradores: `),
            B("LUIZ CLAUDIO GUIMARÃES REHEM DE MATOS"),
            N(`, brasileiro, advogado, inscrito na OAB/SP sob o nº 500.731 e `),
            B("RAFAEL GARCIA PEREIRA"),
            N(`, brasileiro, advogado, inscrito na OAB/SP sob o nº 500.843, com endereço eletrônico: contato.guimaraesgarcia@gmail.com, outorgando-lhes os poderes da `),
            B("cláusula ad judicia,"),
            N(` conforme previsto no art. 105 do Código de Processo Civil, e os demais necessários para o foro em geral e para a defesa dos interesses do outorgante em juízo ou fora dele, em todas as instâncias e graus de jurisdição, inclusive os `),
            B("poderes especiais"),
            N(` para confessar, reconhecer a procedência do pedido, transigir, desistir, renunciar ao direito sobre o qual se funda a ação, receber, dar quitação, firmar compromisso e substabelecer com ou sem reserva de poderes, praticando, enfim, todos os atos judiciais necessários ao fiel desempenho deste mandato.`),
          ],
        }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 800 }, indent: { firstLine: 720 }, children: [N(`São Paulo, ${dia} de ${mes} de ${ano}.`)] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [N("___________________________________________")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [N(d.nome.toUpperCase())] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("Luiz Claudio G. Rehem de Matos. OAB/SP 500.731 e Rafael Garcia Pereira. OAB/SP 500.843.")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [B("Email: contato.guimaraesgarcia@gmail.com.")] }),
      ],
    }],
  });

  download(await Packer.toBlob(doc), `Procuracao_${d.nome.replace(/\s+/g, "_")}.docx`);
}

// ─── Geração DOCX: Contrato ───────────────────────────────────────────────────

async function gerarContrato(d, fin) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, BorderStyle } = await import("https://cdn.skypack.dev/docx@8.5.0");
  const { dia, mes, ano } = dataHoje();
  const B = (t) => new TextRun({ text: t, bold: true, font: "Times New Roman", size: 24 });
  const N = (t) => new TextRun({ text: t, font: "Times New Roman", size: 24 });
  const endereco = `${d.rua}, ${d.numero}, ${d.bairro}, ${d.cidade} - ${d.estado}, CEP ${d.cep}`;

  const borda = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
  const bordas = { top: borda, bottom: borda, left: borda, right: borda };

  const cel = (children, size = 1800) => new TableCell({
    borders: bordas,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    width: { size, type: WidthType.DXA },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children })],
  });

  // Divide parcelas em duas colunas
  const parcelas = fin.parcelas;
  const metade = Math.ceil(parcelas.length / 2);
  const esq = parcelas.slice(0, metade);
  const dir = parcelas.slice(metade);

  const tabelaRows = esq.map((p, i) => {
    const pd = dir[i];
    return new TableRow({
      children: [
        cel([B(`${p.label}:`)], 1700), cel([B(p.valor)], 1400), cel([B(p.data)], 1400),
        cel([], 300),
        ...(pd
          ? [cel([B(`${pd.label}:`)], 1700), cel([B(pd.valor)], 1400), cel([B(pd.data)], 1400)]
          : [cel([], 1700), cel([B("Total:")], 1400), cel([B(fin.valorTotal)], 1400)]
        ),
      ],
    });
  });

  if (parcelas.length % 2 === 0) {
    tabelaRows.push(new TableRow({
      children: [
        cel([], 1700), cel([], 1400), cel([], 1400), cel([], 300),
        cel([], 1700), cel([B("Total:")], 1400), cel([B(fin.valorTotal)], 1400),
      ],
    }));
  }

  const tabela = new Table({
    width: { size: 9300, type: WidthType.DXA },
    columnWidths: [1700, 1400, 1400, 300, 1700, 1400, 1400],
    rows: tabelaRows,
  });

  const P = (children, opts = {}) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 }, ...opts, children });
  const titulo = (t) => P([B(t)], { spacing: { before: 240, after: 100 } });
  const cl = (num, children, isNum = false) => P([B(`Cláusula ${num}${isNum ? "" : "ª"} -\t`), ...children]);

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 } } },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("GUIMARÃES & GARCIA")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B("SOCIEDADE DE ADVOGADOS")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B("CONTRATO DE PRESTAÇÃO DE SERVIÇOS ADVOCATÍCIOS")] }),

        P([B(d.nome.toUpperCase()), N(`, ${d.estado_civil}, ${d.profissao}, inscrito(a) no CPF sob o nº ${d.cpf}, residente e domiciliado na ${endereco}, com endereço eletrônico: ${d.email}, doravante denominado(a) de `), B("CONTRATANTE;")]),
        P([N(" ")]),
        P([N("E, de outro lado, "), B("LUIZ CLAUDIO GUIMARÃES REHEM DE MATOS"), N(", brasileiro, advogado, inscrito na OAB/SP sob o nº 500.731, e "), B("RAFAEL GARCIA PEREIRA,"), N(" brasileiro, advogado, inscrito na OAB/SP sob o nº 500.843, com escritório na Rua Cruzeiro, Barra Funda, São Paulo – SP, CEP 01137-000, e-mail: contato.guimaraesgarcia@gmail.com, doravante denominados "), B("CONTRATADOS"), N(", resolvem, de mútuo e comum acordo, celebrar o presente contrato sob as seguintes cláusulas e condições:")]),

        titulo("CONSTITUIÇÃO DE MANDATO"),
        cl("1", [N("O "), B("CONTRATADO"), N(", independentemente de ordem ou nomeação, desde a assinatura do presente instrumento, fica constituído procurador do "), B("CONTRATANTE"), N(` para prestação de serviços advocatícios, consistentes no acompanhamento processual e na elaboração de petições, a fim de defender seus interesses na forma mais ampla possível, em especial nos autos do processo nº ${fin.numeroProcesso}.`)]),

        titulo("DOS HONORÁRIOS CONTRATUAIS E PAGAMENTOS"),
        P([N(" ")]),
        cl("2", [N("O "), B("CONTRATANTE"), N(" pagará ao "), B("CONTRATADO,"), N(" a título de honorários advocatícios pelos serviços prestados, os seguintes valores:")]),
        P([N(" ")]),

        P([B(`A importância de ${fin.valorTotal} (${fin.valorTotalExtenso}), independentemente de a demanda ser julgada procedente ou improcedente, total ou parcialmente, sendo o pagamento inicial no valor de ${fin.parcelas[0]?.valor || ""} e o saldo remanescente em ${fin.numParcelasRestantes} parcelas de ${fin.valorParcela} cada.`)], { indent: { left: 720 } }),
        P([N(" ")]),

        tabela,
        P([N(" ")]),

        P([B(`O percentual de ${fin.percentualExito} sobre o proveito econômico obtido, em caso de êxito na demanda.`)], { indent: { left: 720 } }),
        P([B("A importância de um salário-mínimo federal, na hipótese de falta injustificada em eventual audiência designada pelo Juízo, caso a referida ausência importe no arquivamento do processo.")], { indent: { left: 360 } }),
        P([B("Em caso de inadimplemento, a multa será de 20% sobre o valor da parcela, acrescida de juros de 2% ao mês e correção monetária pelo INPC (índice nacional de preços ao consumidor).")], { indent: { left: 720 } }),
        P([B("Os pagamentos deverão ocorrer em conta bancária, via PIX, ou por qualquer outro meio disponibilizado pelo CONTRATADO no momento da contratação.")], { indent: { left: 720 } }),

        titulo("DOS HONORÁRIOS DE ÊXITO: INCIDÊNCIA SOBRE O PROVEITO ECONÔMICO"),
        cl("3", [N("O "), B("CONTRATANTE "), N("desde já está ciente que o "), B("percentual acima a ser pago incidirá sobre a totalidade dos reais proveitos econômicos auferidos. "), N("Assim, compreende-se que o cálculo abrangerá valores pecuniários recebidos, bem como valores que foram ou deixarão de ser desembolsados em favor de terceiros, apurados cumulativamente. Inclusive, havendo desconto fiscal compulsório ou quaisquer outras deduções de natureza previdenciária ou tributária sobre o montante, estas serão suportadas exclusivamente pelo(a) "), B("CONTRATANTE"), N(", garantindo-se ao "), B("CONTRATADO"), N(" o recebimento integral do percentual acima sobre o total indicado nos cálculos de liquidação de sentença e/ou acordo, sem quaisquer descontos ou deduções mencionadas anteriormente.")]),

        titulo("DOS HONORÁRIOS SUCUMBENCIAIS"),
        cl("4", [N("O "), B("CONTRATANTE"), N(", com a presente, fica ciente de que os "), B("HONORÁRIOS SUCUMBENCIAIS "), N("constituem verba destinada exclusivamente ao "), B("CONTRATADO"), N(". Por outro lado, é de responsabilidade exclusiva do "), B("CONTRATANTE"), N(" o pagamento de sucumbência em favor da parte contrária.")]),

        titulo("OBRIGAÇÕES DO CONTRATANTE"),
        cl("5", [N("Não se compreendem, nas quantias acima estipuladas, quaisquer despesas judiciais ou extrajudiciais, tais como custas processuais, taxa judiciária, honorários de terceiros, despesas de viagem, despesas com fotocópias de documentos, certidões e demais despesas necessárias.")]),
        cl("6", [N("O "), B("CONTRATANTE"), N(" será obrigado(a) a fornecer o numerário necessário para a satisfação de eventuais despesas, de modo a não interromper o andamento do processo. Não o fazendo, fica o "), B("CONTRATADO"), N(" isento de qualquer responsabilidade pela demora ou interrupção que disso resulte. Ressalta-se que as quantias adiantadas pelo "), B("CONTRATADO"), N(" deverão ser restituídas/pagas sempre que exigidas, a qualquer tempo.")]),

        titulo("DA REVOGAÇÃO, ACORDO OU DESISTÊNCIA"),
        cl("7", [N("O "), B("CONTRATADO "), N("terá direito aos honorários estabelecidos na cláusula 2ª caso o(a) "), B("CONTRATANTE"), N(" revogue o mandato antes da conclusão da causa, transija de qualquer forma com a parte contrária ou impeça o prosseguimento do feito.")]),
        cl("8", [N("Ocorrendo o descumprimento da cláusula 5ª e/ou a efetiva ocorrência de qualquer das hipóteses previstas na cláusula 7ª pelo(a) "), B("CONTRATANTE"), N(", poderá o "), B("CONTRATADO"), N(" considerar o presente contrato violado, reputando-se vencido e imediatamente exigível o saldo e/ou a totalidade dos honorários advocatícios.")]),
        cl("9", [N("Fica estabelecido que, em caso de desistência da ação por parte do(a) "), B("CONTRATANTE"), N(" antes do início dos serviços indicados na cláusula 1ª, serão devidos ao "), B("CONTRATADO"), N(", a título de honorários por assessoria e consultoria jurídica, o valor de R$ 1.000,00 (mil reais). Caso a desistência ocorra após o início dos serviços, estando o feito ainda em Primeira Instância, serão devidos honorários contratuais proporcionais ao trabalho realizado, com base na tabela de honorários mínimos da Ordem dos Advogados do Brasil.")]),

        titulo("DEDUÇÃO DE VALORES POR ALVARÁ JUDICIAL"),
        cl("10", [N("O "), B("CONTRATANTE"), N(" tem ciência de que os valores a serem pagos referente aos honorários advocatícios acima estipulados poderão ser deduzidos do montante a ser levantado por meio de alvará judicial.")], true),

        titulo("INFORMAÇÕES SOBRE ANDAMENTOS PROCESSUAIS"),
        cl("11", [N("O "), B("CONTRATADO"), N(", por meio de seus advogados ou funcionários, prestará informações ao "), B("CONTRATANTE,"), N(" por e-mail ou por escrito (WhatsApp), sempre que solicitado.")], true),
        cl("12", [N("O "), B("CONTRATADO"), N(" prestará informações trimestralmente sobre os andamentos processuais, salvo nos casos em que as informações demandem contato imediato.")], true),

        titulo("DISPOSIÇÕES GERAIS"),
        P([N(" ")]),
        cl("13", [N("O "), B("CONTRATANTE "), N("se compromete a manter seu endereço e telefones atualizados, bem como a fornecer, tempestivamente, todos os meios necessários para o bom desempenho das atividades do "), B("CONTRATADO"), N(", principalmente documentos, provas, informações e assessoramento técnico, quando necessário, isentando-se o "), B("CONTRATADO"), N(" de qualquer responsabilidade por atraso, negligência, caso fortuito ou força maior que impliquem no descumprimento das obrigações processuais dentro do prazo estabelecido em lei.")], true),
        cl("14", [N("É responsabilidade do(a) "), B("CONTRATANTE"), N(" informar ao "), B("CONTRATADO "), N("sobre qualquer fato superveniente que seja relevante para a solução do litígio judicial ou administrativo, bem como sobre o seu interesse em efetuar acordo extrajudicial ou outra forma de composição amigável da lide.")], true),
        cl("15", [N("O presente contrato de honorários, nos termos da lei, constitui, para efeitos legais, título executivo extrajudicial, cuja obrigação é reconhecida pelo(a) "), B("CONTRATANTE"), N(" como líquida, certa e exigível, sendo plenamente exequível.")], true),
        cl("16", [B("O CONTRATANTE DECLARA ESTAR CIENTE DE QUE O PRESENTE CONTRATO É DE MEIO, NA VIA JUDICIAL. (UMA OBRIGAÇÃO DE MEIO É UM VÍNCULO OBRIGACIONAL NO QUAL O CONTRATADO SE COMPROMETE A ENVIDAR SEUS MELHORES ESFORÇOS PARA ALCANÇAR DETERMINADO RESULTADO, SEM GARANTIR QUE O CONSEGUIRÁ.)")], true),
        cl("17", [N("Somente terá validade como recibo de quitação da obrigação ora contraída pelo(a) "), B("CONTRATANTE"), N(" o recibo datado e assinado pelo "), B("CONTRATADO"), N(", no qual seja dada quitação da integralidade dos honorários aqui pactuados.")], true),
        cl("18", [N("Os herdeiros, sucessores ou cessionários da parte "), B("CONTRATANTE "), N("se obrigam, desde já, ao integral cumprimento deste contrato.")], true),
        cl("19", [N("Para dirimir quaisquer controvérsias oriundas do presente instrumento particular, as partes elegem como foro único e privilegiado a comarca de São Paulo – SP, renunciando a qualquer outro.")], true),

        P([N("Por fim, as partes declaram que estão em pleno gozo da capacidade civil e que leram integralmente o presente instrumento, manifestando sua vontade sem qualquer vício de consentimento, reserva mental ou impedimento que possa invalidá-lo. E, por estarem assim justos e contratados, firmam o presente instrumento, com todas as suas cláusulas e subcláusulas, em duas vias de igual teor.")]),
        P([N(`São Paulo, ${dia} de ${mes} de ${ano}.`)], { spacing: { after: 600 } }),

        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("___________________________________________")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B(d.nome.toUpperCase())] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("LUIZ CLAUDIO GUIMARÃES REHEM DE MATOS")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B("OAB/SP nº 500.731")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("RAFAEL GARCIA PEREIRA")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [B("OAB/SP nº 500.843")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [B("Luiz Claudio G. Rehem de Matos. OAB/SP 500.731 e Rafael Garcia Pereira. OAB/SP 500.843.")] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [B("Tel. (11) 98232-7948 ou (11) 95979-3703. Email: contato.guimaraesgarcia@gmail.com.")] }),
      ],
    }],
  });

  download(await Packer.toBlob(doc), `Contrato_${d.nome.replace(/\s+/g, "_")}.docx`);
}

// ─── Estilos ──────────────────────────────────────────────────────────────────

const GOLD = "#c9a84c";
const GOLD_L = "#e8c96b";

const sLabel = { display: "block", fontSize: 11, color: "#8a9bb0", marginBottom: 5, fontFamily: "sans-serif", letterSpacing: 1, textTransform: "uppercase" };
const sInput = (vazio) => ({ width: "100%", background: vazio ? "rgba(180,60,60,0.12)" : "rgba(0,0,0,0.3)", border: vazio ? "1px solid rgba(200,80,80,0.4)" : "1px solid rgba(201,168,76,0.2)", borderRadius: 8, color: "#e8e0d0", fontSize: 14, padding: "9px 12px", fontFamily: "sans-serif", outline: "none", boxSizing: "border-box" });
const sBtn = (dis) => ({ marginTop: 20, width: "100%", padding: "14px 0", background: dis ? "rgba(201,168,76,0.3)" : `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, border: "none", borderRadius: 10, color: dis ? "#8a7a50" : "#1a2940", fontSize: 15, fontWeight: "bold", cursor: dis ? "not-allowed" : "pointer", fontFamily: "sans-serif", letterSpacing: 1 });

function Campo({ label, value, onChange, full, placeholder }) {
  return (
    <div style={{ gridColumn: full ? "1 / -1" : "auto" }}>
      <label style={sLabel}>{label}</label>
      <input value={value || ""} placeholder={placeholder || ""} onChange={(e) => onChange(e.target.value)} style={sInput(!value)} />
    </div>
  );
}

function SecTitle({ children }) {
  return (
    <div style={{ gridColumn: "1 / -1", borderBottom: "1px solid rgba(201,168,76,0.2)", paddingBottom: 6, marginTop: 12, marginBottom: 2 }}>
      <span style={{ fontSize: 11, color: GOLD, letterSpacing: 2, textTransform: "uppercase", fontFamily: "sans-serif" }}>{children}</span>
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [tipo, setTipo] = useState(null);
  const [texto, setTexto] = useState("");
  const [dados, setDados] = useState(null);
  const [fin, setFin] = useState({ numeroProcesso: "", valorTotal: "", valorTotalExtenso: "", valorParcela: "", numParcelasRestantes: "", percentualExito: "10% (dez por cento)", parcelas: parcelasVazias(2) });
  const [etapa, setEtapa] = useState("tipo");
  const [carregando, setCarregando] = useState(false);
  const [erro, setErro] = useState("");

  async function extrairDados() {
    if (!texto.trim()) return;
    setCarregando(true); setErro("");
    try {
      const res = await fetch("/api/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 1000, messages: [{ role: "user", content: EXTRACTION_PROMPT + "\n\n" + texto }] }),
      });
      const data = await res.json();
      const parsed = JSON.parse((data.content?.[0]?.text || "{}").replace(/```json|```/g, "").trim());
      setDados(parsed);
      setEtapa("revisao");
    } catch { setErro("Erro ao extrair dados. Verifique o texto e tente novamente."); }
    finally { setCarregando(false); }
  }

  async function gerarDoc() {
    setEtapa("gerando");
    try {
      if (tipo === "procuracao") await gerarProcuracao(dados);
      else await gerarContrato(dados, fin);
      setEtapa("pronto");
    } catch (e) {
      setErro("Erro ao gerar documento: " + e.message);
      setEtapa(tipo === "contrato" ? "financeiro" : "revisao");
    }
  }

  function resetar() {
    setTipo(null); setTexto(""); setDados(null); setErro("");
    setFin({ numeroProcesso: "", valorTotal: "", valorTotalExtenso: "", valorParcela: "", numParcelasRestantes: "", percentualExito: "10% (dez por cento)", parcelas: parcelasVazias(2) });
    setEtapa("tipo");
  }

  const camposCliente = [
    { key: "nome", label: "Nome completo", full: true },
    { key: "estado_civil", label: "Estado civil" }, { key: "profissao", label: "Profissão" },
    { key: "rg", label: "RG" }, { key: "orgao_expeditor", label: "Órgão expedidor" },
    { key: "cpf", label: "CPF" },
    { key: "rua", label: "Rua", full: true },
    { key: "numero", label: "Número" }, { key: "bairro", label: "Bairro" },
    { key: "cidade", label: "Cidade" }, { key: "estado", label: "Estado (sigla)" }, { key: "cep", label: "CEP" },
    { key: "email", label: "E-mail", full: true },
  ];

  const steps = tipo === "contrato" ? ["Documento","Texto","Dados","Honorários","Pronto"] : ["Documento","Texto","Dados","Pronto"];
  const etapaIdx = { tipo: 0, input: 1, revisao: 2, financeiro: 3, gerando: tipo === "contrato" ? 3 : 2, pronto: tipo === "contrato" ? 4 : 3 }[etapa] ?? 0;

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f1923 0%, #1a2940 100%)", fontFamily: "Georgia, serif", color: "#e8e0d0", padding: "32px 16px" }}>
      <div style={{ maxWidth: 700, margin: "0 auto" }}>

        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{ display: "inline-block", background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", fontSize: 12, letterSpacing: 4, textTransform: "uppercase", marginBottom: 8 }}>
            Guimarães & Garcia · Advogados
          </div>
          <h1 style={{ fontSize: 26, fontWeight: "normal", color: "#f0ead8", margin: "0 0 6px", letterSpacing: 1 }}>Gerador de Documentos</h1>
          <p style={{ color: "#8a9bb0", fontSize: 13, margin: 0, fontFamily: "sans-serif" }}>Procurações e Contratos de Honorários</p>
        </div>

        {/* Steps */}
        {etapa !== "tipo" && (
          <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 28, flexWrap: "wrap" }}>
            {steps.map((s, i) => (
              <div key={s} style={{ padding: "5px 14px", borderRadius: 20, fontSize: 12, fontFamily: "sans-serif", background: i < etapaIdx ? "#1e6b45" : i === etapaIdx ? GOLD : "rgba(255,255,255,0.07)", color: i <= etapaIdx ? "#fff" : "#5a6a7a", transition: "all 0.3s" }}>
                {i < etapaIdx ? "✓ " : ""}{s}
              </div>
            ))}
          </div>
        )}

        {/* Card */}
        <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(201,168,76,0.2)", borderRadius: 16, padding: 32, backdropFilter: "blur(10px)" }}>

          {/* Escolha do tipo */}
          {etapa === "tipo" && (
            <>
              <p style={{ textAlign: "center", color: "#8a9bb0", fontFamily: "sans-serif", fontSize: 14, marginTop: 0, marginBottom: 28 }}>Qual documento deseja gerar?</p>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                {[
                  { id: "procuracao", icon: "📋", title: "Procuração", desc: "Instrumento de mandato para representação judicial" },
                  { id: "contrato", icon: "⚖️", title: "Contrato de Honorários", desc: "Contrato de prestação de serviços advocatícios" },
                ].map((op) => (
                  <button key={op.id} onClick={() => { setTipo(op.id); setEtapa("input"); }}
                    style={{ background: "rgba(0,0,0,0.2)", border: "1px solid rgba(201,168,76,0.25)", borderRadius: 14, padding: "24px 16px", cursor: "pointer", textAlign: "center", color: "#e8e0d0", fontFamily: "Georgia, serif" }}
                    onMouseEnter={e => e.currentTarget.style.borderColor = GOLD}
                    onMouseLeave={e => e.currentTarget.style.borderColor = "rgba(201,168,76,0.25)"}>
                    <div style={{ fontSize: 36, marginBottom: 10 }}>{op.icon}</div>
                    <div style={{ fontSize: 16, fontWeight: "bold", marginBottom: 8, color: GOLD_L }}>{op.title}</div>
                    <div style={{ fontSize: 12, color: "#8a9bb0", fontFamily: "sans-serif", lineHeight: 1.5 }}>{op.desc}</div>
                  </button>
                ))}
              </div>
            </>
          )}

          {/* Texto WhatsApp */}
          {etapa === "input" && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <label style={{ ...sLabel, margin: 0 }}>Texto recebido do cliente (WhatsApp)</label>
                <button onClick={() => setEtapa("tipo")} style={{ background: "none", border: "none", color: "#5a6a7a", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              <textarea value={texto} onChange={(e) => setTexto(e.target.value)}
                placeholder="Cole aqui a mensagem do cliente com os dados pessoais..."
                style={{ width: "100%", minHeight: 200, background: "rgba(0,0,0,0.3)", border: "1px solid rgba(201,168,76,0.25)", borderRadius: 10, color: "#e8e0d0", fontSize: 14, padding: 16, fontFamily: "sans-serif", resize: "vertical", outline: "none", lineHeight: 1.6, boxSizing: "border-box" }} />
              {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 8, fontFamily: "sans-serif" }}>{erro}</p>}
              <button onClick={extrairDados} disabled={carregando || !texto.trim()} style={sBtn(carregando || !texto.trim())}>
                {carregando ? "⚙️  Extraindo dados..." : "Extrair Dados →"}
              </button>
            </>
          )}

          {/* Revisão dados cliente */}
          {etapa === "revisao" && dados && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <span style={{ ...sLabel, margin: 0 }}>Confirme os dados do cliente</span>
                <button onClick={() => setEtapa("input")} style={{ background: "none", border: "none", color: "#5a6a7a", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                {camposCliente.map(({ key, label, full }) => (
                  <Campo key={key} label={label} full={full} value={dados[key]} onChange={(v) => setDados(p => ({ ...p, [key]: v }))} />
                ))}
              </div>
              <button onClick={() => tipo === "contrato" ? setEtapa("financeiro") : gerarDoc()} style={{ ...sBtn(false), marginTop: 24 }}>
                {tipo === "contrato" ? "Próximo: Honorários →" : "📋  Gerar Procuração (.docx)"}
              </button>
            </>
          )}

          {/* Honorários (contrato) */}
          {etapa === "financeiro" && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <span style={{ ...sLabel, margin: 0 }}>Dados financeiros do contrato</span>
                <button onClick={() => setEtapa("revisao")} style={{ background: "none", border: "none", color: "#5a6a7a", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <SecTitle>Processo</SecTitle>
                <Campo label="Número do processo" full value={fin.numeroProcesso} onChange={(v) => setFin(p => ({ ...p, numeroProcesso: v }))} placeholder="0000000-00.0000.0.00.0000" />

                <SecTitle>Honorários</SecTitle>
                <Campo label="Valor total (ex: R$ 10.000,00)" value={fin.valorTotal} onChange={(v) => setFin(p => ({ ...p, valorTotal: v }))} />
                <Campo label="Valor total por extenso" value={fin.valorTotalExtenso} onChange={(v) => setFin(p => ({ ...p, valorTotalExtenso: v }))} placeholder="dez mil reais" />
                <Campo label="Valor de cada parcela" value={fin.valorParcela} onChange={(v) => setFin(p => ({ ...p, valorParcela: v }))} placeholder="R$ 1.000,00" />
                <Campo label="Nº de parcelas restantes" value={fin.numParcelasRestantes} onChange={(v) => setFin(p => ({ ...p, numParcelasRestantes: v }))} placeholder="5" />
                <Campo label="Percentual de êxito" full value={fin.percentualExito} onChange={(v) => setFin(p => ({ ...p, percentualExito: v }))} />

                <SecTitle>Tabela de pagamentos</SecTitle>

                {/* Seletor de linhas */}
                <div style={{ gridColumn: "1 / -1", display: "flex", alignItems: "center", gap: 10, marginBottom: 4, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 12, color: "#8a9bb0", fontFamily: "sans-serif" }}>Linhas na tabela:</span>
                  {[2,3,4,5,6,7,8].map(n => (
                    <button key={n} onClick={() => setFin(p => ({ ...p, parcelas: parcelasVazias(n).map((x, i) => p.parcelas[i] || x) }))}
                      style={{ padding: "4px 12px", borderRadius: 6, border: "1px solid rgba(201,168,76,0.3)", background: fin.parcelas.length === n ? GOLD : "transparent", color: fin.parcelas.length === n ? "#1a2940" : GOLD, cursor: "pointer", fontFamily: "sans-serif", fontSize: 13 }}>
                      {n}
                    </button>
                  ))}
                </div>

                {fin.parcelas.map((p, i) => (
                  <div key={i} style={{ gridColumn: "1 / -1", display: "grid", gridTemplateColumns: "150px 1fr 1fr", gap: 8, alignItems: "end" }}>
                    <div>
                      <label style={sLabel}>Rótulo</label>
                      <input value={p.label} onChange={(e) => setFin(prev => { const ps = [...prev.parcelas]; ps[i] = { ...ps[i], label: e.target.value }; return { ...prev, parcelas: ps }; })} style={sInput(!p.label)} />
                    </div>
                    <div>
                      <label style={sLabel}>Valor</label>
                      <input value={p.valor} placeholder="R$ 0,00" onChange={(e) => setFin(prev => { const ps = [...prev.parcelas]; ps[i] = { ...ps[i], valor: e.target.value }; return { ...prev, parcelas: ps }; })} style={sInput(!p.valor)} />
                    </div>
                    <div>
                      <label style={sLabel}>Data</label>
                      <input value={p.data} placeholder="DD/MM/AAAA" onChange={(e) => setFin(prev => { const ps = [...prev.parcelas]; ps[i] = { ...ps[i], data: e.target.value }; return { ...prev, parcelas: ps }; })} style={sInput(!p.data)} />
                    </div>
                  </div>
                ))}
              </div>

              {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 12, fontFamily: "sans-serif" }}>{erro}</p>}
              <button onClick={gerarDoc} style={{ ...sBtn(false), marginTop: 24 }}>⚖️  Gerar Contrato (.docx)</button>
            </>
          )}

          {/* Gerando */}
          {etapa === "gerando" && (
            <div style={{ textAlign: "center", padding: "32px 0" }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>⚙️</div>
              <p style={{ color: "#8a9bb0", fontFamily: "sans-serif" }}>Gerando documento...</p>
            </div>
          )}

          {/* Pronto */}
          {etapa === "pronto" && (
            <div style={{ textAlign: "center", padding: "20px 0" }}>
              <div style={{ fontSize: 56, marginBottom: 16 }}>✅</div>
              <h2 style={{ color: GOLD_L, fontWeight: "normal", marginBottom: 8 }}>
                {tipo === "procuracao" ? "Procuração gerada!" : "Contrato gerado!"}
              </h2>
              <p style={{ color: "#8a9bb0", fontSize: 14, fontFamily: "sans-serif", marginBottom: 28 }}>
                O arquivo <strong style={{ color: GOLD }}>.docx</strong> foi baixado automaticamente.<br />
                Abra no Word para revisar antes de enviar ao cliente.
              </p>
              <button onClick={resetar} style={{ padding: "12px 32px", background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, border: "none", borderRadius: 10, color: "#1a2940", fontSize: 15, fontWeight: "bold", cursor: "pointer", fontFamily: "sans-serif", letterSpacing: 1 }}>
                + Novo Documento
              </button>
            </div>
          )}
        </div>

        <p style={{ textAlign: "center", color: "#3a4a5a", fontSize: 12, marginTop: 24, fontFamily: "sans-serif" }}>
          Os dados são processados de forma segura e não são armazenados.
        </p>
      </div>
    </div>
  );
}
