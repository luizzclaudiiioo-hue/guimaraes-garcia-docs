import { useState } from "react";

const GOLD = "#9a6e1e";
const GOLD_L = "#b8891f";

const RESUMO_PROMPT = `Você é um assistente jurídico. Resuma a publicação judicial abaixo em até 4 linhas, em português, indicando:
1. O que foi decidido ou determinado;
2. Se há providência a ser tomada pelo advogado e qual;
3. O prazo mencionado, se houver (apenas o que consta no texto; não calcule datas).
Responda em texto simples, sem markdown.

Publicação:`;

const sLabel = { display: "block", fontSize: 11, color: "#555555", marginBottom: 5, fontFamily: "sans-serif", letterSpacing: 1, textTransform: "uppercase" };
const sInput = { width: "100%", boxSizing: "border-box", background: "rgba(0,0,0,0.06)", border: "1px solid rgba(90,122,90,0.3)", borderRadius: 8, color: "#2a2a2a", fontSize: 14, padding: "10px 12px", fontFamily: "sans-serif", outline: "none" };
const sBtn = (dis) => ({ marginTop: 16, width: "100%", padding: "14px 0", background: dis ? "rgba(201,168,76,0.3)" : `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`, border: "none", borderRadius: 10, color: "#1a2940", fontSize: 15, fontWeight: "bold", cursor: dis ? "not-allowed" : "pointer", fontFamily: "sans-serif", letterSpacing: 1 });
const sLink = { background: "none", border: "none", padding: 0, color: GOLD, cursor: "pointer", fontSize: 13, fontFamily: "sans-serif", textDecoration: "underline" };

// localStorage pode falhar (aba anônima, bloqueio do navegador): o app segue sem ele.
function ler(chave, padrao) {
  try { const v = localStorage.getItem(chave); return v === null ? padrao : JSON.parse(v); } catch { return padrao; }
}
function gravar(chave, valor) {
  try { localStorage.setItem(chave, JSON.stringify(valor)); } catch { /* ignora */ }
}

function isoLocal(d) {
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}
function diasAtras(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return isoLocal(d);
}
function dataBR(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || "");
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
}
function textoPlano(html) {
  const doc = new DOMParser().parseFromString(String(html).replace(/<br\s*\/?>/gi, "\n"), "text/html");
  return (doc.body.textContent || "").replace(/\n{3,}/g, "\n\n").trim();
}

export default function Publicacoes({ model }) {
  const [oab, setOab] = useState(() => ler("pub_oab", ""));
  const [uf, setUf] = useState(() => ler("pub_uf", "SP"));
  const [inicio, setInicio] = useState(diasAtras(7));
  const [fim, setFim] = useState(isoLocal(new Date()));
  const [lista, setLista] = useState(null);
  const [lidas, setLidas] = useState(() => ler("pub_lidas", {}));
  const [soNaoLidas, setSoNaoLidas] = useState(true);
  const [abertas, setAbertas] = useState({});
  const [resumos, setResumos] = useState({});
  const [carregando, setCarregando] = useState(false);
  const [erro, setErro] = useState("");

  async function buscar() {
    setCarregando(true); setErro("");
    gravar("pub_oab", oab); gravar("pub_uf", uf);
    try {
      const params = new URLSearchParams({ oab, uf, inicio, fim });
      const res = await fetch(`/api/publicacoes?${params}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erro na consulta.");
      setLista(data.publicacoes.map((p) => ({ ...p, texto: textoPlano(p.texto) })));
    } catch (e) {
      setErro(e.message);
      setLista(null);
    } finally {
      setCarregando(false);
    }
  }

  function alternarLida(id) {
    setLidas((prev) => {
      const novo = { ...prev };
      if (novo[id]) delete novo[id]; else novo[id] = true;
      gravar("pub_lidas", novo);
      return novo;
    });
  }

  async function resumir(p) {
    setResumos((r) => ({ ...r, [p.id]: { carregando: true } }));
    try {
      const res = await fetch("/api/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model, max_tokens: 500, messages: [{ role: "user", content: RESUMO_PROMPT + "\n\n" + p.texto }] }),
      });
      const data = await res.json();
      const texto = data.content?.[0]?.text;
      if (!texto) throw new Error();
      setResumos((r) => ({ ...r, [p.id]: { texto } }));
    } catch {
      setResumos((r) => ({ ...r, [p.id]: { erro: "Não foi possível gerar o resumo." } }));
    }
  }

  const naoLidas = lista ? lista.filter((p) => !lidas[p.id]).length : 0;
  const visiveis = lista ? lista.filter((p) => !soNaoLidas || !lidas[p.id]) : [];
  const podeBuscar = oab.trim() && uf.trim().length === 2 && inicio && fim && !carregando;

  return (
    <>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: 12 }}>
        <div>
          <label style={sLabel}>Nº OAB</label>
          <input value={oab} onChange={(e) => setOab(e.target.value)} placeholder="123456" inputMode="numeric" style={sInput} />
        </div>
        <div>
          <label style={sLabel}>UF</label>
          <input value={uf} onChange={(e) => setUf(e.target.value.toUpperCase().slice(0, 2))} placeholder="SP" style={sInput} />
        </div>
        <div>
          <label style={sLabel}>De</label>
          <input type="date" value={inicio} onChange={(e) => setInicio(e.target.value)} style={sInput} />
        </div>
        <div>
          <label style={sLabel}>Até</label>
          <input type="date" value={fim} onChange={(e) => setFim(e.target.value)} style={sInput} />
        </div>
      </div>
      <button onClick={buscar} disabled={!podeBuscar} style={sBtn(!podeBuscar)}>
        {carregando ? "⚙️  Consultando o DJEN..." : "🔎  Buscar publicações"}
      </button>
      {erro && <p style={{ color: "#e05050", fontSize: 13, marginTop: 12, fontFamily: "sans-serif" }}>{erro}</p>}

      {lista && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, margin: "24px 0 12px", fontFamily: "sans-serif", fontSize: 13, color: "#555555" }}>
            <span>{lista.length} publicação(ões) · <strong style={{ color: GOLD }}>{naoLidas} não lida(s)</strong></span>
            <label style={{ cursor: "pointer" }}>
              <input type="checkbox" checked={soNaoLidas} onChange={(e) => setSoNaoLidas(e.target.checked)} /> Mostrar só não lidas
            </label>
          </div>

          {visiveis.length === 0 && (
            <p style={{ textAlign: "center", color: "#888888", fontFamily: "sans-serif", fontSize: 14, padding: "16px 0" }}>Nenhuma publicação para exibir.</p>
          )}

          {visiveis.map((p) => {
            const lida = !!lidas[p.id];
            const aberta = !!abertas[p.id];
            const resumo = resumos[p.id];
            const longo = p.texto.length > 400;
            return (
              <div key={p.id} style={{ background: "rgba(255,255,255,0.7)", border: `1px solid ${lida ? "rgba(0,0,0,0.08)" : "rgba(154,110,30,0.45)"}`, borderRadius: 12, padding: "16px 16px 12px", marginBottom: 12, opacity: lida ? 0.65 : 1 }}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap", fontFamily: "sans-serif", fontSize: 12, color: "#555555", marginBottom: 6 }}>
                  <span><strong style={{ color: GOLD }}>{dataBR(p.data)}</strong> · {p.tribunal} · {p.tipo}</span>
                  {!lida && <span style={{ color: GOLD, fontWeight: "bold" }}>● NOVA</span>}
                </div>
                <div style={{ fontSize: 15, fontWeight: "bold", marginBottom: 4, wordBreak: "break-word" }}>{p.processo || "Processo não informado"}</div>
                <div style={{ fontFamily: "sans-serif", fontSize: 12, color: "#555555", marginBottom: 10 }}>
                  {[p.classe, p.orgao].filter(Boolean).join(" · ")}
                  {p.partes.length > 0 && <div>Partes: {p.partes.join("; ")}</div>}
                </div>
                <div style={{ fontFamily: "sans-serif", fontSize: 13, lineHeight: 1.55, whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                  {aberta || !longo ? p.texto : p.texto.slice(0, 400) + "…"}
                </div>
                {resumo && (
                  <div style={{ marginTop: 10, padding: "10px 12px", borderRadius: 8, background: "rgba(154,110,30,0.08)", fontFamily: "sans-serif", fontSize: 13, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>
                    {resumo.carregando ? "⚙️  Resumindo..." : resumo.erro ? <span style={{ color: "#e05050" }}>{resumo.erro}</span> : <><strong style={{ color: GOLD }}>Resumo (IA — confira no texto):</strong>{"\n"}{resumo.texto}</>}
                  </div>
                )}
                <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginTop: 12 }}>
                  {longo && <button onClick={() => setAbertas((a) => ({ ...a, [p.id]: !aberta }))} style={sLink}>{aberta ? "Recolher texto" : "Ver texto completo"}</button>}
                  {!resumo?.texto && !resumo?.carregando && <button onClick={() => resumir(p)} style={sLink}>Resumir com IA</button>}
                  {p.link && <a href={p.link} target="_blank" rel="noreferrer" style={{ ...sLink, display: "inline-block" }}>Abrir documento</a>}
                  <button onClick={() => alternarLida(p.id)} style={sLink}>{lida ? "Marcar como não lida" : "✓ Marcar como lida"}</button>
                </div>
              </div>
            );
          })}
        </>
      )}
    </>
  );
}
