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

function parcelasVazias(n) {
  return Array.from({ length: n }, (_, i) => ({
    label: i === 0 ? "Entrada" : `${i}ª parcela`,
    valor: "",
    data: "",
  }));
}

const GOLD = "#c9a84c";
const GOLD_L = "#e8c96a";

const sLabel = { display: "block", fontSize: 11, color: "#555555", marginBottom: 5, fontFamily: "sans-serif", letterSpacing: 1, textTransform: "uppercase" };
const sInput = (vazio) => ({ width: "100%", background: vazio ? "rgba(180,60,60,0.10)" : "rgba(0,0,0,0.06)", border: vazio ? "1px solid rgba(200,80,80,0.5)" : "1px solid rgba(90,122,90,0.3)", borderRadius: 8, color: "#2a2a2a", fontSize: 14, padding: "9px 12px", fontFamily: "sans-serif", outline: "none", boxSizing: "border-box" });
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
    <div style={{ gridColumn: "1 / -1", borderBottom: "1px solid rgba(90,122,90,0.25)", paddingBottom: 6, marginTop: 12, marginBottom: 2 }}>
      <span style={{ fontSize: 11, color: GOLD, letterSpacing: 2, textTransform: "uppercase", fontFamily: "sans-serif" }}>{children}</span>
    </div>
  );
}

export default function App() {
  const [autenticado, setAutenticado] = useState(false);
  const [senha, setSenha] = useState("");
  const [erroSenha, setErroSenha] = useState(false);
  const [tipo, setTipo] = useState(null);
  const [texto, setTexto] = useState("");
  const [dados, setDados] = useState(null);
  const [fin, setFin] = useState({ numeroProcesso: "", valorTotal: "", valorEntrada: "", valorParcela: "", numParcelasRestantes: "", percentualExito: "10% (dez por cento)" });
  const [etapa, setEtapa] = useState("tipo");
  const [carregando, setCarregando] = useState(false);
  const [erro, setErro] = useState("");

  function verificarSenha() {
    const s1 = import.meta.env.VITE_SENHA_1;
    const s2 = import.meta.env.VITE_SENHA_2;
    if (senha === s1 || senha === s2) {
      setAutenticado(true);
      setErroSenha(false);
    } else {
      setErroSenha(true);
      setSenha("");
    }
  }

  if (!autenticado) {
    return (
      <>
        <link href="https://fonts.googleapis.com/css2?family=Libre+Baskerville:wght@700&display=swap" rel="stylesheet" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" />
        <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #3a4a3a 0%, #2e3d2e 100%)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "20px 12px" }}>
          <div style={{ background: "transparent", borderRadius: 16, padding: "40px 32px", width: "100%", maxWidth: 560 }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
              <div style={{ background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", fontSize: "clamp(18px, 4vw, 28px)", letterSpacing: 5, textTransform: "uppercase", fontWeight: "700", fontFamily: "'Libre Baskerville', serif", lineHeight: 1.3, whiteSpace: "nowrap" }}>
                GUIMARÃES & GARCIA
              </div>
              <div style={{ background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", fontSize: "clamp(9px, 1.5vw, 12px)", letterSpacing: 5, textTransform: "uppercase", fontFamily: "'Libre Baskerville', serif", marginTop: 6, opacity: 0.85 }}>
                SOCIEDADE DE ADVOGADOS
              </div>
              <div style={{ width: 60, height: 2, background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, margin: "20px auto 0" }} />
        </div>
            <label style={{ ...sLabel, color: "#c9a84c" }}>Senha de acesso</label>
            <input
              type="password"
              value={senha}
              onChange={(e) => { setSenha(e.target.value); setErroSenha(false); }}
              onKeyDown={(e) => e.key === "Enter" && verificarSenha()}
              placeholder="Digite sua senha"
              style={{ ...sInput(erroSenha), background: "rgba(255,255,255,0.12)", border: erroSenha ? "1px solid #c0392b" : "1px solid rgba(201,168,76,0.4)", color: "#f0e8d0", marginBottom: 4 }}
            />
            {erroSenha && <p style={{ color: "#c0392b", fontSize: 12, fontFamily: "sans-serif", margin: "6px 0 0" }}>Senha incorreta. Tente novamente.</p>}
            <button onClick={verificarSenha} style={sBtn(!senha)}>
              Entrar
            </button>
          </div>
        </div>
      </>
    );
  }

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
      const res = await fetch("/api/gerar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tipo, dados, financeiro: fin }),
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || "Erro ao gerar documento");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const nomesArquivo = {
        procuracao: "Procuracao_",
        contrato: "Contrato_",
        declaracao: "Declaracao_Hipossuficiencia_",
        revogacao: "Revogacao_Mandato_",
        residencia_propria: "Declaracao_Residencia_",
        amaisa: "Declaracao_Amaisa_",
        proprietario: "Declaracao_Proprietario_",
      };
      a.download = `${nomesArquivo[tipo] || "Documento_"}${dados.nome.replace(/\s+/g, "_")}.docx`;
      a.click();
      URL.revokeObjectURL(url);
      setEtapa("pronto");
    } catch (e) {
      setErro("Erro ao gerar documento: " + e.message);
      setEtapa(tipo === "contrato" ? "financeiro" : "revisao");
    }
  }

  function resetar() {
    setTipo(null); setTexto(""); setDados(null); setErro("");
    setFin({ numeroProcesso: "", valorTotal: "", valorEntrada: "", valorParcela: "", numParcelasRestantes: "", percentualExito: "10% (dez por cento)" });
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

  const tipoTemFinanceiro = tipo === "contrato";
  const tipoTemDadosEspeciais = ["amaisa","proprietario","revogacao"].includes(tipo);
  const steps = tipoTemFinanceiro ? ["Documento","Texto","Dados","Honorários","Pronto"] : tipoTemDadosEspeciais ? ["Documento","Texto","Dados","Extras","Pronto"] : ["Documento","Texto","Dados","Pronto"];
  const etapaIdx = { tipo: 0, input: 1, revisao: 2, financeiro: 3, extras: 3, gerando: tipoTemFinanceiro ? 3 : tipoTemDadosEspeciais ? 3 : 2, pronto: tipoTemFinanceiro ? 4 : tipoTemDadosEspeciais ? 4 : 3 }[etapa] ?? 0;

  return (
    <>
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #3a4a3a 0%, #2e3d2e 100%)", fontFamily: "Georgia, serif", color: "#2a2a2a", padding: "20px 12px" }}>
      <div style={{ maxWidth: 700, margin: "0 auto" }}>

        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{ display: "inline-block", background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", fontSize: "clamp(16px, 5vw, 22px)", letterSpacing: 3, textTransform: "uppercase", fontWeight: "700", fontFamily: "'Libre Baskerville', serif", lineHeight: 1.3 }}>
            GUIMARÃES & GARCIA
          </div>
          <div style={{ display: "block", background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", fontSize: 11, letterSpacing: 3, textTransform: "uppercase", fontFamily: "'Libre Baskerville', serif", marginTop: 4, opacity: 0.85 }}>
            SOCIEDADE DE ADVOGADOS
          </div>
        </div>

        {etapa !== "tipo" && (
          <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 28, flexWrap: "wrap" }}>
            {steps.map((s, i) => (
              <div key={s} style={{ padding: "5px 14px", borderRadius: 20, fontSize: 12, fontFamily: "sans-serif", background: i < etapaIdx ? "#4a7a4a" : i === etapaIdx ? GOLD : "rgba(255,255,255,0.12)", color: i <= etapaIdx ? "#fff" : "#aaaaaa", transition: "all 0.3s" }}>
                {i < etapaIdx ? "✓ " : ""}{s}
              </div>
            ))}
          </div>
        )}

        <div style={{ background: "#d4dbd4", border: "1px solid rgba(90,122,90,0.35)", borderRadius: 16, padding: "24px 20px", backdropFilter: "blur(10px)", boxShadow: "0 2px 20px rgba(0,0,0,0.25)" }}>

          {etapa === "tipo" && (
            <>
              <p style={{ textAlign: "center", color: "#555555", fontFamily: "sans-serif", fontSize: 14, marginTop: 0, marginBottom: 28 }}>Qual documento deseja gerar?</p>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 12 }}>
                {[
                  { id: "procuracao", icon: "📋", title: "Procuração", desc: "Instrumento de mandato para representação judicial" },
                  { id: "contrato", icon: "⚖️", title: "Contrato de Honorários", desc: "Contrato de prestação de serviços advocatícios" },
                  { id: "declaracao", icon: "📝", title: "Declaração de Justiça Gratuita", desc: "Declaração de hipossuficiência econômica" },
                  { id: "revogacao", icon: "🚫", title: "Revogação de Mandato", desc: "Revogação de procuração outorgada anteriormente" },
                  { id: "residencia_propria", icon: "🏠", title: "Declaração de Residência Própria", desc: "Declaração de residência do proprietário do imóvel" },
                  { id: "amaisa", icon: "💑", title: "Declaração de Amásia", desc: "Declaração de união estável" },
                  { id: "proprietario", icon: "🔑", title: "Declaração Residencial (Proprietário)", desc: "Declaração do proprietário ao locatário" },
                ].map((op) => (
                  <button key={op.id} onClick={() => { setTipo(op.id); setEtapa("input"); }}
                    style={{ background: "rgba(255,255,255,0.7)", border: "1px solid rgba(90,122,90,0.25)", borderRadius: 14, padding: "24px 16px", cursor: "pointer", textAlign: "center", color: "#2a2a2a", fontFamily: "Georgia, serif" }}
                    onMouseEnter={e => e.currentTarget.style.borderColor = GOLD}
                    onMouseLeave={e => e.currentTarget.style.borderColor = "rgba(201,168,76,0.25)"}>
                    <div style={{ fontSize: 36, marginBottom: 10 }}>{op.icon}</div>
                    <div style={{ fontSize: 16, fontWeight: "bold", marginBottom: 8, color: GOLD }}>{op.title}</div>
                    <div style={{ fontSize: 12, color: "#555555", fontFamily: "sans-serif", lineHeight: 1.5 }}>{op.desc}</div>
                  </button>
                ))}
              </div>
            </>
          )}

          {etapa === "input" && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <label style={{ ...sLabel, margin: 0 }}>Texto recebido do cliente (WhatsApp)</label>
                <button onClick={() => setEtapa("tipo")} style={{ background: "none", border: "none", color: "#666666", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              <textarea value={texto} onChange={(e) => setTexto(e.target.value)}
                placeholder="Cole aqui a mensagem do cliente com os dados pessoais..."
                style={{ width: "100%", minHeight: 200, background: "rgba(0,0,0,0.06)", border: "1px solid rgba(90,122,90,0.3)", borderRadius: 10, color: "#2a2a2a", fontSize: 14, padding: 16, fontFamily: "sans-serif", resize: "vertical", outline: "none", lineHeight: 1.6, boxSizing: "border-box" }} />
              {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 8, fontFamily: "sans-serif" }}>{erro}</p>}
              <button onClick={extrairDados} disabled={carregando || !texto.trim()} style={sBtn(carregando || !texto.trim())}>
                {carregando ? "⚙️  Extraindo dados..." : "Extrair Dados →"}
              </button>
            </>
          )}

          {etapa === "revisao" && dados && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <span style={{ ...sLabel, margin: 0 }}>Confirme os dados do cliente</span>
                <button onClick={() => setEtapa("input")} style={{ background: "none", border: "none", color: "#666666", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 12 }}>
                {camposCliente.map(({ key, label, full }) => (
                  <Campo key={key} label={label} full={full} value={dados[key]} onChange={(v) => setDados(p => ({ ...p, [key]: v }))} />
                ))}
              </div>
              <button onClick={() => tipo === "contrato" ? setEtapa("financeiro") : tipoTemDadosEspeciais ? setEtapa("extras") : gerarDoc()} style={{ ...sBtn(false), marginTop: 24 }}>
                {tipoTemFinanceiro ? "Próximo: Honorários →" : tipoTemDadosEspeciais ? "Próximo: Dados Extras →" : tipo === "declaracao" ? "📝  Gerar Declaração (.docx)" : "📋  Gerar Procuração (.docx)"}
              </button>
            </>
          )}

          {etapa === "financeiro" && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <span style={{ ...sLabel, margin: 0 }}>Dados financeiros do contrato</span>
                <button onClick={() => setEtapa("revisao")} style={{ background: "none", border: "none", color: "#666666", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 12 }}>
                <SecTitle>Processo</SecTitle>
                <Campo label="Número do processo" full value={fin.numeroProcesso} onChange={(v) => setFin(p => ({ ...p, numeroProcesso: v }))} placeholder="0000000-00.0000.0.00.0000" />

                <SecTitle>Honorários</SecTitle>
                <Campo label="Valor total (ex: R$ 10.000,00)" value={fin.valorTotal} onChange={(v) => setFin(p => ({ ...p, valorTotal: v }))} />
                <Campo label="Valor da entrada" value={fin.valorEntrada} onChange={(v) => setFin(p => ({ ...p, valorEntrada: v }))} placeholder="R$ 0,00" />
                <Campo label="Valor de cada parcela" value={fin.valorParcela} onChange={(v) => setFin(p => ({ ...p, valorParcela: v }))} placeholder="R$ 1.000,00" />
                <Campo label="Nº de parcelas restantes" value={fin.numParcelasRestantes} onChange={(v) => setFin(p => ({ ...p, numParcelasRestantes: v }))} placeholder="5" />
                <Campo label="Percentual de êxito" full value={fin.percentualExito} onChange={(v) => setFin(p => ({ ...p, percentualExito: v }))} />


              </div>
              {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 12, fontFamily: "sans-serif" }}>{erro}</p>}
              <button onClick={gerarDoc} style={{ ...sBtn(false), marginTop: 24 }}>⚖️  Gerar Contrato (.docx)</button>
            </>
          )}

          {etapa === "extras" && (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <span style={{ ...sLabel, margin: 0 }}>Dados adicionais</span>
                <button onClick={() => setEtapa("revisao")} style={{ background: "none", border: "none", color: "#666666", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>← Voltar</button>
              </div>
              {tipo === "amaisa" && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 12 }}>
                  <SecTitle>Dados da União Estável</SecTitle>
                  <Campo label="Data de início da união" value={fin.dataUniao || ""} onChange={(v) => setFin(p => ({ ...p, dataUniao: v }))} placeholder="dd/mm/aaaa" />
                  <SecTitle>Dados do Preso</SecTitle>
                  <Campo label="Nome completo do preso" full value={fin.nomePreso || ""} onChange={(v) => setFin(p => ({ ...p, nomePreso: v }))} />
                  <Campo label="RG do preso" value={fin.rgPreso || ""} onChange={(v) => setFin(p => ({ ...p, rgPreso: v }))} />
                  <Campo label="Órgão expedidor (ex: SSP/SP)" value={fin.orgaoPreso || ""} onChange={(v) => setFin(p => ({ ...p, orgaoPreso: v }))} placeholder="SSP/SP" />
                  <Campo label="CPF do preso" value={fin.cpfPreso || ""} onChange={(v) => setFin(p => ({ ...p, cpfPreso: v }))} />
                  <Campo label="Local de detenção" full value={fin.localDetencao || ""} onChange={(v) => setFin(p => ({ ...p, localDetencao: v }))} placeholder="CDP de Chácara Belém I, São Paulo" />
                  <Campo label="Matrícula" value={fin.matricula || ""} onChange={(v) => setFin(p => ({ ...p, matricula: v }))} />
                </div>
              )}
              {tipo === "proprietario" && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 12 }}>
                  <SecTitle>Endereço do Imóvel Alugado</SecTitle>
                  <Campo label="Endereço completo do imóvel" full value={fin.enderecoImovel || ""} onChange={(v) => setFin(p => ({ ...p, enderecoImovel: v }))} placeholder="Rua Exemplo, 123, Bairro, São Paulo – SP, CEP 00000-000" />
                  <SecTitle>Dados do Locatário (Cliente)</SecTitle>
                  <Campo label="Nome do locatário" full value={fin.nomeLocatario || ""} onChange={(v) => setFin(p => ({ ...p, nomeLocatario: v }))} />
                  <Campo label="CPF do locatário" value={fin.cpfLocatario || ""} onChange={(v) => setFin(p => ({ ...p, cpfLocatario: v }))} />
                </div>
              )}
              {tipo === "revogacao" && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 12 }}>
                  <SecTitle>Dados da Revogação</SecTitle>
                  <Campo label="Número do processo" full value={fin.numeroProcesso || ""} onChange={(v) => setFin(p => ({ ...p, numeroProcesso: v }))} placeholder="0000000-00.0000.0.00.0000" />
                </div>
              )}
              {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 12, fontFamily: "sans-serif" }}>{erro}</p>}
              <button onClick={gerarDoc} style={{ ...sBtn(false), marginTop: 24 }}>📄  Gerar Documento (.docx)</button>
            </>
          )}

          {etapa === "gerando" && (
            <div style={{ textAlign: "center", padding: "32px 0" }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>⚙️</div>
              <p style={{ color: "#555555", fontFamily: "sans-serif" }}>Gerando documento...</p>
            </div>
          )}

          {etapa === "pronto" && (
            <div style={{ textAlign: "center", padding: "20px 0" }}>
              <div style={{ fontSize: 56, marginBottom: 16 }}>✅</div>
              <h2 style={{ color: GOLD_L, fontWeight: "normal", marginBottom: 8 }}>
                {tipo === "procuracao" ? "Procuração gerada!" : tipo === "contrato" ? "Contrato gerado!" : "Documento gerado!"}
              </h2>
              <p style={{ color: "#555555", fontSize: 14, fontFamily: "sans-serif", marginBottom: 28 }}>
                O arquivo <strong style={{ color: GOLD }}>.docx</strong> foi baixado automaticamente.<br />
                Abra no Word para revisar antes de enviar ao cliente.
              </p>
              <button onClick={resetar} style={{ padding: "12px 32px", background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, border: "none", borderRadius: 10, color: "#1a2940", fontSize: 15, fontWeight: "bold", cursor: "pointer", fontFamily: "sans-serif", letterSpacing: 1 }}>
                + Novo Documento
              </button>
            </div>
          )}
        </div>

        <p style={{ textAlign: "center", color: "#888888", fontSize: 12, marginTop: 24, fontFamily: "sans-serif" }}>
          Os dados são processados de forma segura e não são armazenados.
        </p>
      </div>
    </div>
    </>
  );
}
