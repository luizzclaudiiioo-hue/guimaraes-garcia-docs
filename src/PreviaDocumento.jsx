import { useEffect, useRef, useState } from "react";
import { renderAsync } from "docx-preview";
import fonte400 from "@fontsource/archivo-narrow/files/archivo-narrow-latin-400-normal.woff2?url";
import fonte400i from "@fontsource/archivo-narrow/files/archivo-narrow-latin-400-italic.woff2?url";
import fonte700 from "@fontsource/archivo-narrow/files/archivo-narrow-latin-700-normal.woff2?url";
import fonte700i from "@fontsource/archivo-narrow/files/archivo-narrow-latin-700-italic.woff2?url";

const GOLD = "#9a6e1e";
const GOLD_L = "#b8891f";
const LARGURA_PAGINA = 794; // A4 (21cm) em px

// Os modelos usam "Arial Nova Cond" (fonte do Office). Se o aparelho não tiver,
// tenta Arial Narrow e, por fim, a Archivo Narrow que vai junto com o app.
// Cada variação aponta para o arquivo dela ("Arial Nova Cond Bold" etc.): apontar o
// negrito para a fonte normal instalada fazia o negrito sumir no PDF.
function cssFontes() {
  const abs = (u) => new URL(u, window.location.href).href;
  const face = (url, peso, estilo, sufixo) => {
    const locais = sufixo
      ? [`Arial Nova Cond ${sufixo}`, `ArialNovaCond-${sufixo.replace(" ", "")}`, `ArialNova-Cond${sufixo.replace(" ", "")}`, `Arial Narrow ${sufixo}`, `ArialNarrow-${sufixo.replace(" ", "")}`]
      : ["Arial Nova Cond", "ArialNovaCond", "ArialNova-Cond", "Arial Narrow", "ArialNarrow"];
    return `@font-face { font-family: "Arial Nova Cond"; font-weight: ${peso}; font-style: ${estilo};
    src: ${locais.map((n) => `local("${n}")`).join(", ")}, url("${abs(url)}") format("woff2"); }`;
  };
  return [face(fonte400, 400, "normal", ""), face(fonte400i, 400, "italic", "Italic"), face(fonte700, 700, "normal", "Bold"), face(fonte700i, 700, "italic", "Bold Italic")].join("\n");
}

const OPCOES_RENDER = { className: "docx", inWrapper: true, breakPages: true, renderHeaders: true, renderFooters: true };

const px = (v) => parseFloat(v) || 0;

// Monta o HTML para impressão: o texto corre livre entre as páginas, com margens do Word;
// cabeçalho e rodapé ficam fixos (repetem em toda página) e os espaçadores de thead/tfoot
// reservam a área deles em cada página.
function htmlImpressao(container, titulo) {
  const secoes = Array.from(container.querySelectorAll("section.docx"));
  const sec = secoes[0];
  const cs = getComputedStyle(sec);
  const [padT, padR, padB, padL] = [cs.paddingTop, cs.paddingRight, cs.paddingBottom, cs.paddingLeft].map(px);
  const cab = sec.querySelector(":scope > header");
  const rod = sec.querySelector(":scope > footer");
  const cabTopo = cab ? padT + px(getComputedStyle(cab).marginTop) : 0;
  const rodBase = rod ? padB + px(getComputedStyle(rod).marginBottom) : 0;
  const estilos = Array.from(container.querySelectorAll("style")).map((el) => el.outerHTML).join("\n");
  const artigos = secoes.map((s, i) => {
    const art = s.querySelector(":scope > article");
    return `<div class="${i > 0 ? "quebra" : ""}">${art ? art.innerHTML : ""}</div>`;
  }).join("");

  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${titulo.replace(/[<>&"]/g, "")}</title>
    <style>${cssFontes()}</style>${estilos}
    <style>
      @page { size: A4; margin: 0; }
      html, body { margin: 0; padding: 0; background: #fff; }
      .docx-wrapper { background: #fff !important; padding: 0 !important; display: block !important; }
      .docx-wrapper > section.docx { padding: 0 !important; margin: 0 !important; width: auto !important; min-height: 0 !important; box-shadow: none !important; display: block !important; }
      .cab, .rod { position: fixed; left: ${padL}px; right: ${padR}px; margin: 0 !important; min-height: 0 !important; }
      .cab { top: ${cabTopo}px; } .rod { bottom: ${rodBase}px; }
      table.pg { border-collapse: collapse; width: 100%; table-layout: fixed; }
      table.pg > * > tr > td { padding: 0 ${padR}px 0 ${padL}px; }
      .quebra { break-before: page; page-break-before: always; }
      .aviso { font: 14px sans-serif; background: #fdf3dc; color: #5a4210; padding: 12px 16px; border-bottom: 1px solid #e4cf9c; position: relative; z-index: 1; }
      @media screen { .cab, .rod { position: static; } }
      @media print { .aviso { display: none; } }
    </style></head>
    <body><div class="aviso">Na janela de impressão, escolha <b>Salvar como PDF</b> como destino.
    No iPhone: toque em Compartilhar → Imprimir → Compartilhar novamente → Salvar em Arquivos.</div>
    <div class="docx-wrapper"><section class="docx">
      ${cab ? `<header class="cab">${cab.innerHTML}</header>` : ""}
      ${rod ? `<footer class="rod">${rod.innerHTML}</footer>` : ""}
      <table class="pg">
        <thead><tr><td><div id="esp-topo" data-min="${padT}" data-inicio="${cabTopo}"></div></td></tr></thead>
        <tfoot><tr><td><div id="esp-base" data-min="${padB}" data-inicio="${rodBase}"></div></td></tr></tfoot>
        <tbody><tr><td>${artigos}</td></tr></tbody>
      </table>
    </section></div></body></html>`;
}

// Reserva no topo e na base de cada página a área ocupada pelo cabeçalho e pelo rodapé.
// A altura é medida já na aba de impressão, com o cabeçalho na posição fixa.
function ajustarEspacos(doc) {
  [["esp-topo", ".cab"], ["esp-base", ".rod"]].forEach(([id, sel]) => {
    const esp = doc.getElementById(id);
    const el = doc.querySelector(sel);
    let altura = 0;
    if (el) {
      el.style.position = "fixed";
      altura = el.offsetHeight;
      el.style.position = "";
    }
    const total = Math.max(Number(esp.dataset.min), el ? Number(esp.dataset.inicio) + altura : 0);
    esp.style.height = `${total}px`;
  });
}

function baixar(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nome;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function PreviaDocumento({ blob, nomeArquivo, titulo, onVoltar, onNovo }) {
  const areaRef = useRef(null);
  const docRef = useRef(null);
  const [zoom, setZoom] = useState(1);
  const [status, setStatus] = useState("carregando");

  useEffect(() => {
    let ativo = true;
    setStatus("carregando");
    renderAsync(blob, docRef.current, null, { ...OPCOES_RENDER, ignoreLastRenderedPageBreak: false })
      .then(() => ativo && setStatus("ok"))
      .catch(() => ativo && setStatus("erro"));
    return () => { ativo = false; };
  }, [blob]);

  // Reduz a página A4 para caber na largura da tela (celular)
  useEffect(() => {
    const ajustar = () => {
      const largura = areaRef.current?.clientWidth || LARGURA_PAGINA;
      setZoom(Math.min(1, largura / (LARGURA_PAGINA + 24)));
    };
    ajustar();
    window.addEventListener("resize", ajustar);
    return () => window.removeEventListener("resize", ajustar);
  }, []);

  // Abre o documento numa aba própria e chama a impressão: "Salvar como PDF"
  // gera um PDF com texto pesquisável, aceito pelos sistemas dos tribunais.
  async function baixarPdf() {
    // A aba precisa ser aberta já no clique, senão o navegador a bloqueia
    const janela = window.open("", "_blank");
    if (!janela) {
      alert("O navegador bloqueou a nova aba. Permita pop-ups para este site e tente de novo.");
      return;
    }
    janela.document.write('<p style="font-family:sans-serif;padding:16px">Preparando o PDF...</p>');

    // Renderiza de novo, fora da tela e sem zoom, só com as quebras de página explícitas
    const oculto = document.createElement("div");
    oculto.style.cssText = "position:absolute;left:-10000px;top:0;width:900px";
    document.body.appendChild(oculto);
    try {
      await renderAsync(blob, oculto, null, { ...OPCOES_RENDER, ignoreLastRenderedPageBreak: true });
      await document.fonts.ready;
      const html = htmlImpressao(oculto, nomeArquivo);
      janela.document.open();
      janela.document.write(html);
      janela.document.close();
      await janela.document.fonts.ready;
      ajustarEspacos(janela.document);
      setTimeout(() => { janela.focus(); janela.print(); }, 300);
    } catch {
      janela.close();
      alert("Não foi possível gerar o PDF. Baixe o arquivo Word.");
    } finally {
      oculto.remove();
    }
  }

  const sBtnAcao = (primario) => ({ flex: "1 1 160px", padding: "13px 0", borderRadius: 10, cursor: "pointer", fontFamily: "sans-serif", fontSize: 14, fontWeight: "bold", letterSpacing: 0.5,
    border: primario ? "none" : `1px solid ${GOLD}`, background: primario ? `linear-gradient(135deg, ${GOLD}, ${GOLD_L})` : "rgba(255,255,255,0.7)", color: primario ? "#1a2940" : GOLD });

  return (
    <>
      <style>{cssFontes()}</style>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 8 }}>
        <span style={{ fontSize: 11, color: "#555555", fontFamily: "sans-serif", letterSpacing: 1, textTransform: "uppercase" }}>Prévia — {titulo}</span>
        <button onClick={onVoltar} style={{ background: "none", border: "none", color: "#666666", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif", whiteSpace: "nowrap" }}>← Corrigir dados</button>
      </div>

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
        <button onClick={() => baixar(blob, nomeArquivo + ".docx")} style={sBtnAcao(true)}>📄 Baixar Word</button>
        <button onClick={baixarPdf} disabled={status !== "ok"} style={{ ...sBtnAcao(false), opacity: status === "ok" ? 1 : 0.5 }}>📑 Baixar PDF</button>
      </div>

      <div ref={areaRef} style={{ background: "#9aa39a", borderRadius: 10, padding: 12, maxHeight: "70vh", overflow: "auto" }}>
        {status === "carregando" && <p style={{ textAlign: "center", color: "#fff", fontFamily: "sans-serif" }}>Montando a prévia...</p>}
        {status === "erro" && <p style={{ textAlign: "center", color: "#fff", fontFamily: "sans-serif" }}>Não foi possível mostrar a prévia. Você ainda pode baixar o arquivo Word.</p>}
        <div ref={docRef} className="previa-docx" style={{ zoom }} />
      </div>
      <style>{`.previa-docx .docx-wrapper { background: transparent; padding: 0; } .previa-docx section.docx { margin-bottom: 12px; }`}</style>

      <p style={{ color: "#666666", fontSize: 12, fontFamily: "sans-serif", lineHeight: 1.5, margin: "12px 0 20px" }}>
        A prévia e o PDF podem ter pequenas diferenças de fonte e quebra de página em relação ao Word. Confira antes de protocolar.
      </p>
      <button onClick={onNovo} style={{ width: "100%", padding: "12px 0", background: "none", border: "1px solid rgba(0,0,0,0.15)", borderRadius: 10, color: "#555555", fontSize: 14, cursor: "pointer", fontFamily: "sans-serif" }}>
        + Novo Documento
      </button>
    </>
  );
}
