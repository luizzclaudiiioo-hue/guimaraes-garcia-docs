import { useState, useEffect } from "react";
import { supabase } from "./lib/supabase";
import Auth from "./pages/Auth";
import Dashboard from "./pages/Dashboard";

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

export default function App() {
  const [session, setSession] = useState(undefined);
  const [tela, setTela] = useState("dashboard");
  const [tipo, setTipo] = useState(null);
  const [texto, setTexto] = useState("");
  const [dados, setDados] = useState(null);
  const [fin, setFin] = useState({ numeroProcesso: "", valorTotal: "", valorEntrada: "", valorParcela: "", numParcelasRestantes: "", percentualExito: "10% (dez por cento)" });
  const [etapa, setEtapa] = useState("tipo");
  const [carregando, setCarregando] = useState(false);
  const [erro, setErro] = useState("");

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });
    return () => subscription.unsubscribe();
  }, []);

  async function debitarCredito(tipoDoc, nomeCliente) {
    const user = session?.user;
    if (!user) return;
    const { data: cred } = await supabase.from("credits").select("saldo").eq("user_id", user.id).single();
    if (!cred || cred.saldo <= 0) throw new Error("Sem créditos disponíveis.");
    await supabase.from("credits").update({ saldo: cred.saldo - 1, updated_at: new Date().toISOString() }).eq("user_id", user.id);
    await supabase.from("documents").insert({
      user_id: user.id,
      tipo: tipoDoc,
      titulo: tipoDoc === "procuracao" ? `Procuração — ${nomeCliente}` : `Contrato de Honorários — ${nomeCliente}`,
    });
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
      await debitarCredito(tipo, dados.nome);
      const res = await fetch("/api/gerar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tipo, dados, financeiro: fin }),
      });
      if (!res.ok) { const err = await res.json(); throw new Error(err.error || "Erro ao gerar documento"); }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = tipo === "procuracao" ? `Procuracao_${dados.nome.replace(/\s+/g, "_")}.docx` : `Contrato_${dados.nome.replace(/\s+/g, "_")}.docx`;
      a.click();
      URL.revokeObjectURL(url);
      setEtapa("pronto");
    } catch (e) {
      setErro("Erro: " + e.message);
      setEtapa(tipo === "contrato" ? "financeiro" : "revisao");
    }
  }

  function resetar() {
    setTipo(null); setTexto(""); setDados(null); setErro("");
    setFin({ numeroProcesso: "", valorTotal: "", valorEntrada: "", valorParcela: "", numParcelasRestantes: "", percentualExito: "10% (dez por cento)" });
    setEtapa("tipo");
    setTela("dashboard");
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

  const steps = tipo === "contrato" ? ["Documento", "Texto", "Dados", "Honorários", "Pronto"] : ["Documento", "Texto", "Dados", "Pronto"];
  const etapaIdx = { tipo: 0, input: 1, revisao: 2, financeiro: 3, gerando: tipo === "contrato" ? 3 : 2, pronto: tipo === "contrato" ? 4 : 3 }[etapa] ?? 0;

  if (session === undefined) {
    return (
      <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f1923 0%, #1a2940 100%)", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: "#5a6a7a", fontFamily: "sans-serif", fontSize: 14 }}>Carregando...</div>
      </div>
    );
  }

  if (!session) return <Auth />;

  if (tela === "dashboard") {
    return <Dashboard user={session.user} onNovoDocumento={() => { setTela("gerador"); setEtapa("tipo"); }} />;
  }

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f1923 0%, #1a2940 100%)", fontFamily: "Georgia, serif", color: "#e8e0d0", padding: "32px 16px" }}>
      <div style={{ maxWidth: 700, margin: "0 auto" }}>

        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{ display: "inline-block", background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", fontSize: 12, letterSpacing: 4, textTransform: "uppercase", marginBottom: 8 }}>
            Guimarães & Garcia · Advogados
          </div>
          <h1 style={{ fontSize: 26, fontWeight: "normal", color: "#f0ead8", margin: "0 0 6px", letterSpacing: 1 }}>Gerador de Documentos</h1>
          <p style={{ color: "#8a9bb0", fontSize: 13, margin: 0, fontFamily: "sans-serif" }}>Procurações e Contratos de Honorários</p>
        </div>

        <div style={{ textAlign: "left", marginBottom: 16 }}>
          <button onClick={() => setTela("dashboard")} style={{ background: "none", border: "none", color: "#5a6a7a", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif" }}>
            ← Voltar ao painel
          </button>
        </div>

        {etapa !== "tipo" && (
          <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 28, flexWrap: "wrap" }}>
            {steps.map((s, i) => (
              <div key={s} style={{ padding: "5px 14px", borderRadius: 20, fontSize: 12, fontFamily: "sans-serif", background: i < etapaIdx ? "#1e6b45" : i === etapaIdx ? GOLD : "rgba(255,255,255,0.07)", color: i <= etapaIdx ? "#fff" : "#5a6a7a", transition: "all 0.3s" }}>
                {i < etapaIdx ? "✓ " : ""}{s}
              </div>
            ))}
          </div>
        )}

        <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(201,168,76,0.2)", borderRadius: 16, padding: 32, backdropFilter: "blur(10px)" }}>

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
                <Campo label="Valor da entrada" value={fin.valorEntrada} onChange={(v) => setFin(p => ({ ...p, valorEntrada: v }))} placeholder="R$ 0,00" />
                <Campo label="Valor de cada parcela" value={fin.valorParcela} onChange={(v) => setFin(p => ({ ...p, valorParcela: v }))} placeholder="R$ 1.000,00" />
                <Campo label="Nº de parcelas restantes" value={fin.numParcelasRestantes} onChange={(v) => setFin(p => ({ ...p, numParcelasRestantes: v }))} placeholder="5" />
                <Campo label="Percentual de êxito" full value={fin.percentualExito} onChange={(v) => setFin(p => ({ ...p, percentualExito: v }))} />
              </div>
              {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 12, fontFamily: "sans-serif" }}>{erro}</p>}
              <button onClick={gerarDoc} style={{ ...sBtn(false), marginTop: 24 }}>⚖️  Gerar Contrato (.docx)</button>
            </>
          )}

          {etapa === "gerando" && (
            <div style={{ textAlign: "center", padding: "32px 0" }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>⚙️</div>
              <p style={{ color: "#8a9bb0", fontFamily: "sans-serif" }}>Gerando documento...</p>
            </div>
          )}

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
