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
  const [carregando, setCarregando] = useState(
