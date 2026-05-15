import { useState, useEffect } from "react";
import { supabase } from "../lib/supabase";

const GOLD = "#c9a84c";
const GOLD_L = "#e8c96b";

export default function Dashboard({ user, onNovoDocumento }) {
  const [perfil, setPerfil] = useState(null);
  const [creditos, setCreditos] = useState(null);
  const [documentos, setDocumentos] = useState([]);
  const [carregando, setCarregando] = useState(true);

  useEffect(() => {
    carregarDados();
  }, [user]);

  async function carregarDados() {
    setCarregando(true);
    const [{ data: p }, { data: c }, { data: d }] = await Promise.all([
      supabase.from("profiles").select("*").eq("id", user.id).single(),
      supabase.from("credits").select("*").eq("user_id", user.id).single(),
      supabase.from("documents").select("*").eq("user_id", user.id).order("created_at", { ascending: false }).limit(20),
    ]);
    setPerfil(p);
    setCreditos(c);
    setDocumentos(d || []);
    setCarregando(false);
  }

  async function sair() {
    await supabase.auth.signOut();
  }

  const saldo = creditos?.saldo ?? 0;

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #0f1923 0%, #1a2940 100%)",
      fontFamily: "Georgia, serif",
      color: "#e8e0d0",
      padding: "32px 16px",
    }}>
      <div style={{ maxWidth: 700, margin: "0 auto" }}>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 32 }}>
          <div>
            <div style={{
              background: `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`,
              WebkitBackgroundClip: "text",
              WebkitTextFillColor: "transparent",
              fontSize: 11,
              letterSpacing: 4,
              textTransform: "uppercase",
              marginBottom: 6,
            }}>
              Guimarães & Garcia · Advogados
            </div>
            <h1 style={{ fontSize: 22, fontWeight: "normal", color: "#f0ead8", margin: 0 }}>
              Olá, {perfil?.nome?.split(" ")[0] || "Advogado"}
            </h1>
            {perfil?.oab && (
              <p style={{ color: "#5a6a7a", fontSize: 12, margin: "4px 0 0", fontFamily: "sans-serif" }}>OAB {perfil.oab}</p>
            )}
          </div>
          <button
            onClick={sair}
            style={{ background: "none", border: "1px solid rgba(90,106,122,0.4)", borderRadius: 8, color: "#5a6a7a", padding: "8px 16px", cursor: "pointer", fontFamily: "sans-serif", fontSize: 12 }}
          >
            Sair
          </button>
        </div>

        <div style={{
          background: "rgba(255,255,255,0.04)",
          border: "1px solid rgba(201,168,76,0.2)",
          borderRadius: 16,
          padding: 24,
          marginBottom: 24,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          flexWrap: "wrap",
          gap: 16,
        }}>
          <div>
            <div style={{ fontSize: 11, color: "#8a9bb0", letterSpacing: 2, textTransform: "uppercase", fontFamily: "sans-serif", marginBottom: 6 }}>
              Créditos disponíveis
            </div>
            <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
              <span style={{
                fontSize: 40,
                fontWeight: "bold",
                color: saldo === 0 ? "#e05050" : GOLD_L,
                fontFamily: "sans-serif",
              }}>
                {carregando ? "—" : saldo}
              </span>
              <span style={{ color: "#5a6a7a", fontSize: 13, fontFamily: "sans-serif" }}>
                {saldo === 1 ? "crédito" : "créditos"}
              </span>
            </div>
            <p style={{ color: "#5a6a7a", fontSize: 12, margin: "4px 0 0", fontFamily: "sans-serif" }}>
              1 crédito = 1 documento gerado
            </p>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 10, alignItems: "flex-end" }}>
            <button
              onClick={onNovoDocumento}
              disabled={saldo === 0}
              style={{
                padding: "12px 24px",
                background: saldo === 0 ? "rgba(201,168,76,0.15)" : `linear-gradient(135deg, ${GOLD}, ${GOLD_L})`,
                border: "none",
                borderRadius: 10,
                color: saldo === 0 ? "#8a7a50" : "#1a2940",
                fontSize: 14,
                fontWeight: "bold",
