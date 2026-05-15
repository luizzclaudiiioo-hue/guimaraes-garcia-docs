import { useState } from "react";
import { supabase } from "../lib/supabase";

const GOLD = "#c9a84c";
const GOLD_L = "#e8c96b";

const sInput = {
  width: "100%",
  background: "rgba(0,0,0,0.3)",
  border: "1px solid rgba(201,168,76,0.25)",
  borderRadius: 8,
  color: "#e8e0d0",
  fontSize: 14,
  padding: "11px 14px",
  fontFamily: "sans-serif",
  outline: "none",
  boxSizing: "border-box",
  marginTop: 6,
};

const sLabel = {
  display: "block",
  fontSize: 11,
  color: "#8a9bb0",
  letterSpacing: 1,
  textTransform: "uppercase",
  fontFamily: "sans-serif",
};

export default function Auth() {
  const [modo, setModo] = useState("login");
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [nome, setNome] = useState("");
  const [oab, setOab] = useState("");
  const [carregando, setCarregando] = useState(false);
  const [erro, setErro] = useState("");
  const [msg, setMsg] = useState("");

  async function handleLogin(e) {
    e.preventDefault();
    setCarregando(true); setErro("");
    const { error } = await supabase.auth.signInWithPassword({ email, password: senha });
    if (error) setErro("E-mail ou senha incorretos.");
    setCarregando(false);
  }

  async function handleCadastro(e) {
    e.preventDefault();
    setCarregando(true); setErro("");
    const { error } = await supabase.auth.signUp({
      email,
      password: senha,
      options: { data: { nome, oab } },
    });
    if (error) {
      setErro(error.message);
    } else {
      setModo("confirmacao");
      setMsg(`Enviamos um link de confirmação para ${email}. Verifique sua caixa de entrada.`);
    }
    setCarregando(false);
  }

  async function handleReset(e) {
    e.preventDefault();
    setCarregando(true); setErro("");
    const { error } = await supabase.auth.resetPasswordForEmail(email);
    if (error) setErro(error.message);
    else setMsg("Link de redefinição enviado para seu e-mail.");
    setCarregando(false);
  }

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #0f1923 0%, #1a2940 100%)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      padding: "32px 16px",
      fontFamily: "Georgia, serif",
      color: "#e8e0d0",
    }}>
      <div style={{ width: "100%", maxWidth: 420 }}>

        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{
            display: "inline-block",
            background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`,
            WebkitBackgroundClip: "text",
            WebkitTextFillColor: "transparent",
            fontSize: 11,
            letterSpacing: 4,
            textTransform: "uppercase",
            marginBottom: 10,
          }}>
            Guimarães & Garcia · Advogados
          </div>
          <h1 style={{ fontSize: 24, fontWeight: "normal", color: "#f0ead8", margin: "0 0 6px", letterSpacing: 1 }}>
            {modo === "cadastro" ? "Criar conta" : modo === "confirmacao" ? "Confirme seu e-mail" : "Entrar"}
          </h1>
          <p style={{ color: "#8a9bb0", fontSize: 13, margin: 0, fontFamily: "sans-serif" }}>
            {modo === "cadastro" ? "Crie sua conta e ganhe 3 créditos grátis" : "Plataforma de Documentos Jurídicos"}
          </p>
        </div>

        <div style={{
          background: "rgba(255,255,255,0.04)",
          border: "1px solid rgba(201,168,76,0.2)",
          borderRadius: 16,
          padding: 32,
          backdropFilter: "blur(10px)",
        }}>

          {modo === "confirmacao" ? (
            <div style={{ textAlign: "center" }}>
              <div style={{ fontSize: 52, marginBottom: 16 }}>📬</div>
              <p style={{ color: "#8a9bb0", fontFamily: "sans-serif", fontSize: 14, lineHeight: 1.7 }}>{msg}</p>
              <button
                onClick={() => { setModo("login"); setMsg(""); }}
                style={{ marginTop: 20, background: "none", border: `1px solid ${GOLD}`, borderRadius: 8, color: GOLD, padding: "10px 24px", cursor: "pointer", fontFamily: "sans-serif", fontSize: 13 }}
              >
                Voltar ao login
              </button>
            </div>
          ) : (
            <form onSubmit={modo === "login" ? handleLogin : handleCadastro} style={{ display: "flex", flexDirection: "column", gap: 16 }}>

              {modo === "cadastro" && (
                <>
                  <div>
                    <label style={sLabel}>Nome completo</label>
                    <input value={nome} onChange={e => setNome(e.target.value)} required placeholder="Dr. João da Silva" style={sInput} />
                  </div>
                  <div>
                    <label style={sLabel}>OAB (opcional)</label>
                    <input value={oab} onChange={e => setOab(e.target.value)} placeholder="SP 123456" style={sInput} />
                  </div>
                </>
              )}

              <div>
                <label style={sLabel}>E-mail</label>
                <input type="email" value={email} onChange={e => setEmail(e.target.value)} required placeholder="seu@email.com" style={sInput} />
              </div>

              <div>
                <label style={sLabel}>Senha</label>
                <input type="password" value={senha} onChange={e => setSenha(e.target.value)} required placeholder="••••••••" style={sInput} />
              </div>

              {erro && (
                <p style={{ color: "#e05050", fontSize: 13, margin: 0, fontFamily: "sans-serif" }}>{erro}</p>
              )}
              {msg && (
                <p style={{ color: "#50c878", fontSize: 13, margin: 0, fontFamily: "sans-serif" }}>{msg}</p>
              )}

              <button
                type="submit"
                disabled={carregando}
                style={{
                  marginTop: 4,
                  padding: "13px 0",
                  background: carregando ? "rgba(201,168,76,0.3)" : `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`,
                  border: "none",
                  borderRadius: 10,
                  color: carregando ? "#8a7a50" : "#1a2940",
                  fontSize: 15,
                  fontWeight: "bold",
                  cursor: carregando ? "not-allowed" : "pointer",
                  fontFamily: "sans-serif",
                  letterSpacing: 1,
                }}
              >
                {carregando ? "Aguarde..." : modo === "log
