import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { flushSync } from "react-dom";
import fonte400 from "@fontsource/archivo-narrow/files/archivo-narrow-latin-400-normal.woff2?url";
import fonte400i from "@fontsource/archivo-narrow/files/archivo-narrow-latin-400-italic.woff2?url";
import fonte700 from "@fontsource/archivo-narrow/files/archivo-narrow-latin-700-normal.woff2?url";
import fonte700i from "@fontsource/archivo-narrow/files/archivo-narrow-latin-700-italic.woff2?url";
import { CM, PAGINA_PADRAO, dimensoes, docxParaHtml, htmlParaDocx } from "./docxConversao.js";
import * as T from "./textoTabelas.js";
import { ESTILOS_PADRAO, cssEstilos, estilosDe } from "./textoEstilos.js";
import * as R from "./textoRevisao.js";
import { TIPOS, arquivoLembrado, baixar, escolherArquivo, gravarNoArquivo, lembrarArquivo, recursosNaoSuportados, salvarComo, suportaArquivoLocal } from "./arquivoLocal.js";

const GOLD = "#9a6e1e";
const CHAVE = "editorTexto.docs";

// localStorage pode falhar (aba anônima, bloqueio do navegador): o editor segue sem ele.
function ler(chave, padrao) {
  try { const v = localStorage.getItem(chave); return v === null ? padrao : JSON.parse(v); } catch { return padrao; }
}
function gravar(chave, valor) {
  try { localStorage.setItem(chave, JSON.stringify(valor)); return true; } catch { return false; }
}

function novoDoc(nome = "Documento sem título", extra = {}) {
  return {
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    nome, html: "<p><br></p>", cabecalho: "", rodape: "", numPag: "", pagina: PAGINA_PADRAO,
    atualizado: Date.now(), ...extra,
  };
}

// Montador de páginas (paged.js) carregado na janela de impressão/PDF.
const pagedUrl = new URL("../node_modules/pagedjs/dist/paged.polyfill.min.js", import.meta.url).href;
// Arquivos do app em texto/base64 (para embutir na página do PDF), guardados após a 1ª vez.
const cacheRecursos = new Map();
function recursoTexto(url) {
  if (!cacheRecursos.has(url)) cacheRecursos.set(url, fetch(url).then((r) => { if (!r.ok) throw new Error("recurso indisponível"); return r.text(); }));
  return cacheRecursos.get(url);
}
function recursoDataUrl(url) {
  const chave = "data:" + url;
  if (!cacheRecursos.has(chave)) cacheRecursos.set(chave, fetch(url).then((r) => r.blob()).then((b) => new Promise((ok, erro) => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.onerror = erro; fr.readAsDataURL(b); })));
  return cacheRecursos.get(chave);
}
// Fontes de mesmas medidas para o PDF gerado no servidor, que não tem Times, Arial etc.:
// entram só as das famílias usadas no documento.
const ARQ_SUBSTITUTAS = {
    "tinos-400-normal": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-400-normal.woff2", import.meta.url).href,
    "tinos-400-italic": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-400-italic.woff2", import.meta.url).href,
    "tinos-700-normal": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-700-normal.woff2", import.meta.url).href,
    "tinos-700-italic": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-700-italic.woff2", import.meta.url).href,
    "arimo-400-normal": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-400-normal.woff2", import.meta.url).href,
    "arimo-400-italic": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-400-italic.woff2", import.meta.url).href,
    "arimo-700-normal": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-700-normal.woff2", import.meta.url).href,
    "arimo-700-italic": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-700-italic.woff2", import.meta.url).href,
    "carlito-400-normal": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-400-normal.woff2", import.meta.url).href,
    "carlito-400-italic": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-400-italic.woff2", import.meta.url).href,
    "carlito-700-normal": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-700-normal.woff2", import.meta.url).href,
    "carlito-700-italic": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-700-italic.woff2", import.meta.url).href,
    "cousine-400-normal": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-400-normal.woff2", import.meta.url).href,
    "cousine-400-italic": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-400-italic.woff2", import.meta.url).href,
    "cousine-700-normal": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-700-normal.woff2", import.meta.url).href,
    "cousine-700-italic": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-700-italic.woff2", import.meta.url).href,
};
const SUBSTITUTAS = [
  [["Times New Roman", "Times", "Georgia", "Cambria", "Book Antiqua", "Garamond"], "tinos"],
  [["Arial", "Helvetica", "Verdana", "Tahoma", "Segoe UI"], "arimo"],
  [["Calibri"], "carlito"],
  [["Courier New", "Courier"], "cousine"],
];
async function fontesSubstitutas(html) {
  let css = "";
  for (const [familias, arq] of SUBSTITUTAS) {
    const usadas = familias.filter((f) => new RegExp(`["'\\s,:=]${f}["'\\s,;]`, "i").test(html));
    for (const f of usadas) {
      for (const [peso, estilo] of [[400, "normal"], [400, "italic"], [700, "normal"], [700, "italic"]]) {
        const dado = await recursoDataUrl(ARQ_SUBSTITUTAS[`${arq}-${peso}-${estilo}`]);
        css += `@font-face { font-family: "${f}"; font-weight: ${peso}; font-style: ${estilo}; src: local("${f}"), url("${dado}") format("woff2"); }\n`;
      }
    }
  }
  return css;
}
async function fontesEmbutidas() {
  let css = CSS_FONTES;
  for (const url of [...new Set([...CSS_FONTES.matchAll(/url\("([^"]+)"\)/g)].map((m) => m[1]))]) {
    css = css.split(`url("${url}")`).join(`url("${await recursoDataUrl(new URL(url, location.href).href)}")`);
  }
  return css;
}
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// Remove scripts e atributos de evento de HTML vindo de fora.
function sanitizar(html) {
  const corpo = new DOMParser().parseFromString(html, "text/html").body;
  corpo.querySelectorAll("script,style,iframe,object,embed,link,meta").forEach((n) => n.remove());
  corpo.querySelectorAll("*").forEach((n) => [...n.attributes].forEach((a) => { if (/^on/i.test(a.name) || /^\s*javascript:/i.test(a.value)) n.removeAttribute(a.name); }));
  return corpo.innerHTML;
}

const FONTES = ["Times New Roman", "Arial", "Arial Narrow", "Calibri", "Georgia", "Verdana", "Courier New"];
const TAMANHOS = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36];

// Arial Narrow (e a Arial Nova Cond dos modelos) nem sempre estão instaladas: na tela,
// usa a do computador se houver e, se não, a Archivo Narrow que vai junto com o app.
// No .docx o nome gravado continua sendo o original, e o Word usa a fonte verdadeira.
// Cada variação aponta para o arquivo certo (ex.: "Arial Narrow Bold"); apontar o negrito
// para a Arial Narrow normal fazia o negrito sumir em computadores com a fonte instalada.
const VARIACOES = [[fonte400, 400, "normal", ""], [fonte400i, 400, "italic", "Italic"], [fonte700, 700, "normal", "Bold"], [fonte700i, 700, "italic", "Bold Italic"]];
const CSS_FONTES = [["Arial Narrow", "ArialNarrow"], ["Arial Nova Cond", "ArialNova-Cond"]].flatMap(([nome, ps]) => VARIACOES.map(([url, peso, estilo, sufixo]) => {
  const locais = [
    local(sufixo ? `${nome} ${sufixo}` : nome), local(sufixo ? `${ps}-${sufixo.replace(" ", "")}` : ps),
    local(sufixo ? `Arial Narrow ${sufixo}` : "Arial Narrow"), local(sufixo ? `ArialNarrow-${sufixo.replace(" ", "")}` : "ArialNarrow"),
  ];
  return `@font-face { font-family: "${nome}"; font-weight: ${peso}; font-style: ${estilo}; src: ${[...new Set(locais)].join(", ")}, url("${url}") format("woff2"); }`;
})).join("\n");
function local(n) { return `local("${n}")`; }

const FONTE_UI = "'Segoe UI', system-ui, -apple-system, sans-serif";
const COR_TITULO = "#8a6a2c"; // dourado do app, em tom mais leve
const COR_ABAS = "#ece6d8";
const COR_FAIXA = "#f8f5ee";
const COR_STATUS = "#5a6458";
const COR_FUNDO_ESCURO = "#6b7468"; // em volta da página: mais escuro que a folha, sem pesar
const ABAS = [["arquivo", "Arquivo"], ["inicio", "Página Inicial"], ["inserir", "Inserir"], ["layout", "Layout"], ["referencias", "Referências"], ["revisao", "Revisão"], ["exibir", "Exibir"]];
const SIMBOLOS = ["§", "º", "ª", "–", "—", "…", "“", "”", "«", "»", "©", "®", "™", "°", "½", "€", "•", "¶"];
const sRapido = { background: "transparent", border: 0, color: "#fff", fontSize: 15, cursor: "pointer", padding: "2px 6px", borderRadius: 4 };
const sColuna = { display: "flex", flexDirection: "column", gap: 3, justifyContent: "center" };
const sLinha = { display: "flex", gap: 2, alignItems: "center" };
const sCor = { display: "inline-flex", alignItems: "center", gap: 2, height: 26, padding: "0 5px", border: "1px solid transparent", borderRadius: 3, cursor: "pointer", fontSize: 13, fontWeight: 600 };
const sInputCor = { width: 16, height: 16, border: 0, padding: 0, background: "none", cursor: "pointer" };
const sRotulo = { fontSize: 12, color: "#555", width: 92 };
const sZoomBt = { background: "none", border: "1px solid rgba(255,255,255,0.25)", color: "inherit", borderRadius: 3, width: 22, height: 20, cursor: "pointer", fontSize: 12, padding: 0 };

// Grupo da faixa de opções, com a legenda embaixo (como no Word).
function Grupo({ titulo, children }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", borderRight: "1px solid #ddd6c6", padding: "0 8px", margin: "0 0 2px" }}>
      <div style={{ display: "flex", gap: 4, alignItems: "flex-start", flex: 1 }}>{children}</div>
      <div style={{ fontSize: 11, color: "#7a7468", textAlign: "center", paddingTop: 3, whiteSpace: "nowrap" }}>{titulo}</div>
    </div>
  );
}

// Botão da faixa: grande (ícone em cima, texto embaixo) ou pequeno (ícone + texto na linha).
function Btn({ icone, rotulo, grande, ativo, onClick, title }) {
  const [sobre, setSobre] = useState(false);
  const fundo = ativo ? "#f0e2c2" : sobre ? "#efe8d8" : "transparent";
  const estilo = grande
    ? { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-start", gap: 2, minWidth: 54, maxWidth: 84, height: 66, padding: "4px 4px", fontSize: 12, lineHeight: 1.15 }
    : { display: "flex", alignItems: "center", gap: 5, height: 26, minWidth: 26, padding: "0 6px", fontSize: 13 };
  return (
    <button onClick={onClick} title={title || (typeof rotulo === "string" ? rotulo : undefined)} onMouseEnter={() => setSobre(true)} onMouseLeave={() => setSobre(false)}
      style={{ ...estilo, background: fundo, border: `1px solid ${ativo ? "#d8bf86" : sobre ? "#e2d8c2" : "transparent"}`, borderRadius: 3, cursor: "pointer", color: "#2a2a2a", fontFamily: FONTE_UI, justifyContent: grande ? "flex-start" : "center", textAlign: "center" }}>
      <span style={{ fontSize: grande ? 22 : 14, lineHeight: 1, color: ativo ? GOLD : "#3a3a3a" }}>{icone}</span>
      {rotulo && <span style={{ whiteSpace: grande ? "normal" : "nowrap" }}>{rotulo}</span>}
    </button>
  );
}

// Botão com menu suspenso.
function Menu({ icone, rotulo, grande, title, itens }) {
  const [aberto, setAberto] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!aberto) return;
    const fora = (e) => { if (!ref.current?.contains(e.target)) setAberto(false); };
    window.addEventListener("mousedown", fora);
    return () => window.removeEventListener("mousedown", fora);
  }, [aberto]);
  return (
    <span ref={ref} style={{ position: "relative" }}>
      <Btn grande={grande} icone={icone} rotulo={rotulo ? `${rotulo} ▾` : "▾"} ativo={aberto} title={title} onClick={() => setAberto((x) => !x)} />
      {aberto && (
        <div style={{ position: "absolute", top: grande ? 68 : 28, left: 0, zIndex: 30, background: "#fff", border: "1px solid #d0c8b6", borderRadius: 4, boxShadow: "0 4px 16px rgba(0,0,0,0.18)", padding: "4px 0", minWidth: 220 }}>
          {itens.map(([t, f]) => (
            <button key={t} onMouseDown={(e) => e.preventDefault()} onClick={() => { setAberto(false); f(); }}
              onMouseEnter={(e) => { e.currentTarget.style.background = "#f4eee2"; }} onMouseLeave={(e) => { e.currentTarget.style.background = "none"; }}
              style={{ display: "block", width: "100%", textAlign: "left", border: 0, background: "none", padding: "6px 12px", fontSize: 13, cursor: "pointer", fontFamily: FONTE_UI, whiteSpace: "nowrap" }}>{t}</button>
          ))}
        </div>
      )}
    </span>
  );
}

function ContagemPalavras({ editor, paginas, fechar }) {
  const texto = (editor?.innerText || "").replace(/\u200b/g, "");
  const palavras = (texto.match(/\S+/g) || []).length;
  const comEsp = texto.replace(/\n/g, "").length;
  const semEsp = texto.replace(/\s/g, "").length;
  const paragrafos = [...(editor?.querySelectorAll("p,h1,h2,h3,li,blockquote") || [])].filter((p) => p.textContent.trim()).length;
  const linha = (a, b) => <tr><td style={{ padding: "4px 16px 4px 0" }}>{a}</td><td style={{ textAlign: "right", fontWeight: 600 }}>{b.toLocaleString("pt-BR")}</td></tr>;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 80, display: "flex", alignItems: "center", justifyContent: "center" }} onMouseDown={fechar}>
      <div style={{ background: "#fff", borderRadius: 8, padding: 18, minWidth: 280, fontFamily: FONTE_UI, fontSize: 13, boxShadow: "0 8px 30px rgba(0,0,0,0.3)" }} onMouseDown={(e) => e.stopPropagation()}>
        <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 10 }}>Contar palavras</div>
        <table><tbody>
          {linha("Páginas (estimativa)", paginas)}{linha("Palavras", palavras)}{linha("Caracteres (sem espaços)", semEsp)}{linha("Caracteres (com espaços)", comEsp)}{linha("Parágrafos", paragrafos)}
        </tbody></table>
        <div style={{ textAlign: "right", marginTop: 12 }}><button onClick={fechar} style={{ padding: "4px 16px", border: "1px solid #ccc", borderRadius: 4, background: "#fff", cursor: "pointer" }}>Fechar</button></div>
      </div>
    </div>
  );
}

const sBotao = (ativo) => ({ minWidth: 32, height: 30, padding: "0 8px", border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: ativo ? "rgba(154,110,30,0.2)" : "#fff", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif", color: "#2a2a2a", whiteSpace: "nowrap" });
const sSelect = { height: 30, border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: "#fff", fontSize: 13, fontFamily: "sans-serif", padding: "0 4px" };
const sSep = { width: 1, height: 22, background: "rgba(0,0,0,0.15)", margin: "0 4px" };
const sCampo = { ...sSelect, padding: "0 8px" };

// Aparência do texto na página; o .docx gerado usa os mesmos padrões.
const CSS_PAGINA = `
.ed-texto { color: #111; outline: none; }
.ed-texto del.rev { display: none; } .ed-texto ins.rev { text-decoration: none; }
.ed-texto sup.nota { color: #9a6e1e; cursor: pointer; font-weight: bold; padding: 0 1px; }
.ed-notas { font-size: 10pt; line-height: 1.2; }
.ed-notas p { margin: 0 0 2pt; text-indent: 0; }
.sumario { border: 1px dashed #d8cfb8; border-radius: 4px; padding: 6pt 8pt; margin: 0 0 8pt; cursor: default; user-select: none; }
.sumario .sum-titulo { font-weight: bold; font-size: 1.15em; margin-bottom: 4pt; text-align: center; }
.sumario .sum-item { display: flex; gap: 4px; margin: 0; text-indent: 0; cursor: pointer; }
.sumario .sum-item span.pontos { flex: 1; border-bottom: 1px dotted #999; margin-bottom: 4px; }
.sumario .sum-n2 { padding-left: 1.2em; } .sumario .sum-n3 { padding-left: 2.4em; }
.ed-texto table { border-collapse: collapse; } .ed-texto td, .ed-texto th { vertical-align: top; }
.ed-texto hr { border: 0; border-top: 1px dashed #bbb; margin: 8pt 0; }
.ed-texto span.tab { display: inline-block; white-space: pre; text-indent: 0; }
.ed-texto img { max-width: 100%; }
.ed-area { font-size: 10pt; line-height: 1.3; min-height: 1.2em; outline: none; border: 1px dashed transparent; border-radius: 4px; padding: 2px 4px; }
.ed-area:hover, .ed-area:focus { border-color: #d8cfb8; }
.ed-area:empty::before { content: attr(data-vazio); color: #b9b2a3; }
.ed-area p { margin: 0; }
`;

// Formatos de numeração em vários níveis (valem também na impressão).
const CSS_LISTAS = `
.ed-texto ol[data-num="juridico"] li::marker, .ed-texto ol[data-num="juridico"] ol li::marker { content: counters(list-item, ".") ". "; }
.ed-texto ol[data-num="alfa"] > li::marker { content: counter(list-item) ". "; }
.ed-texto ol[data-num="alfa"] ol > li::marker { content: counter(list-item, lower-alpha) ") "; }
.ed-texto ol[data-num="alfa"] ol ol > li::marker { content: counter(list-item, lower-roman) ") "; }
.ed-texto ol[data-num="romano"] > li::marker { content: counter(list-item, upper-roman) ". "; }
.ed-texto ol[data-num="romano"] ol > li::marker { content: counter(list-item, upper-alpha) ". "; }
.ed-texto ol[data-num="romano"] ol ol > li::marker { content: counter(list-item) ". "; }
.ed-texto ol[data-num="letras"] > li::marker { content: counter(list-item, lower-alpha) ") "; }
.ed-texto ol[data-num="letras"] ol > li::marker { content: counter(list-item, lower-roman) ") "; }
.ed-texto ol ol, .ed-texto ol ul, .ed-texto ul ol, .ed-texto ul ul { margin-top: 0; margin-bottom: 0; }
`;

// Só na tela (não vai para a impressão): seleção de células e linhas-guia.
const CSS_TELA = `
.ed-texto td[data-sel], .ed-texto th[data-sel] { background-image: linear-gradient(rgba(60,110,200,0.22), rgba(60,110,200,0.22)) !important; }
.ed-texto td, .ed-texto th { outline: 1px dotted #d6d0c2; outline-offset: -1px; }
.ed-texto table { margin-bottom: 6pt; }
.ed-texto del.rev { display: inline; text-decoration: line-through; }
.ed-texto ins.rev { text-decoration: underline; }
.ed-texto.sem-marcas del.rev { display: none; }
.ed-texto.sem-marcas ins.rev { text-decoration: none; color: inherit !important; }
.ed-texto span.coment { background: rgba(255, 213, 79, 0.35); border-bottom: 2px solid #f0b400; }
.ed-texto.com-pincel { cursor: copy; }
::highlight(busca) { background-color: #ffd54f; color: inherit; }
.ed-texto hr.secao { border: 0; border-top: 3px double #9a6e1e; margin: 12pt 0 14pt; overflow: visible; position: relative; }
.ed-texto hr.secao::after { position: absolute; left: 50%; top: -9px; transform: translateX(-50%); background: #fff; padding: 0 8px; font: 11px sans-serif; color: #9a6e1e; white-space: nowrap; }
.ed-texto hr.secao[data-orient="paisagem"]::after { content: "Quebra de seção · páginas seguintes em paisagem"; }
.ed-texto hr.secao[data-orient="retrato"]::after { content: "Quebra de seção · páginas seguintes em retrato"; }
`;

export default function EditorTexto() {
  const [docs, setDocs] = useState(() => {
    const salvos = ler(CHAVE, null);
    return salvos?.length ? salvos : [novoDoc()];
  });
  const [atualId, setAtualId] = useState(() => docs[0].id);
  const [estado, setEstado] = useState({});
  const [msg, setMsg] = useState("");
  const [quebras, setQuebras] = useState([]); // posições (px) onde as páginas terminam
  const [pagAtual, setPagAtual] = useState(1);
  const [busca, setBusca] = useState(null); // { termo, novo, caso, info }
  const achado = useRef(null); // ocorrência atual da busca (destacada sem tirar o foco da caixa)
  const [painelPagina, setPainelPagina] = useState(false);
  const [menu, setMenu] = useState(null); // menu do botão direito: { x, y, td }
  const [seletorTabela, setSeletorTabela] = useState(null); // { l, c } sob o mouse
  const selCelulas = useRef(null); // { tabela, inicio, ret }
  const [imgSel, setImgSel] = useState(null); // imagem selecionada (elemento)
  const [painelEstilos, setPainelEstilos] = useState(false);
  const [abaFaixa, setAbaFaixa] = useState("inicio");
  const [pincel, setPincel] = useState(null); // formatação copiada pelo pincel
  const [zoom, setZoom] = useState(100);
  const zoomRef = useRef(1);
  const [contagem, setContagem] = useState(false);
  const [mostrarQuebras, setMostrarQuebras] = useState(true);
  const [mostrarGrade, setMostrarGrade] = useState(true);
  const [mostrarRegua, setMostrarRegua] = useState(() => ler("editorTexto.regua", true));
  const [fundoEscuro, setFundoEscuro] = useState(() => ler("editorTexto.fundoEscuro", true));
  const areaRef = useRef(null);
  const [autor, setAutor] = useState(() => ler("editorTexto.autor", ""));
  const [mostrarMarcas, setMostrarMarcas] = useState(true);
  const [comAtivo, setComAtivo] = useState(null); // comentário destacado
  const [comNovo, setComNovo] = useState(null); // comentário recém-criado (abre para escrever)
  const pagRef = useRef(null);
  const desfazerRev = useRef([]); // desfazer próprio enquanto o controle de alterações está ligado
  const [, setTick] = useState(0); // redesenha a moldura da imagem
  const editor = useRef(null);
  const cab = useRef(null);
  const rod = useRef(null);
  const arquivo = useRef(null);
  const arquivoImg = useRef(null);
  const campoBusca = useRef(null);
  const arquivos = useRef(new Map()); // id do documento → arquivo no computador
  const [, setVersaoArquivos] = useState(0);
  const atual = docs.find((x) => x.id === atualId) || docs[0];
  const pagina = atual.pagina || PAGINA_PADRAO;
  const estilos = estilosDe(atual);
  zoomRef.current = zoom / 100;
  const cssDoc = cssEstilos(estilos);
  // Notas de rodapé na ordem em que aparecem no texto.
  const autoresRev = [...new Set([...atual.html.matchAll(/data-autor="([^"]*)"/g)].map((x) => x[1]))];
  const comentarios = atual.comentarios || {};
  const resolvidos = Object.entries(comentarios).filter(([, c]) => c.resolvido).map(([id]) => id);
  const cssRevisao = R.cssAutores(autoresRev)
    + resolvidos.map((id) => `.ed-texto span.coment[data-c="${id}"] { background: none; border-bottom: 1px dotted #ccc; }`).join("")
    + (comAtivo ? `.ed-texto span.coment[data-c="${comAtivo}"] { background: rgba(255, 193, 7, 0.75); }` : "");
  const ordemNotas = [...new Set([...atual.html.matchAll(/data-nota="([^"]+)"/g)].map((x) => x[1]))];
  const [largCm, altCm] = dimensoes(pagina);
  const m = pagina.margens;
  const alturaUtil = (altCm - m.sup - m.inf) * CM;
  // Seções na outra orientação (ex.: páginas em paisagem num documento em retrato).
  const orientBase = pagina.orient === "paisagem" ? "paisagem" : "retrato";
  const [largOutra, altOutra] = dimensoes({ ...pagina, orient: orientBase === "paisagem" ? "retrato" : "paisagem" });
  const deltaSec = largOutra - largCm; // quanto a área de texto cresce (ou encolhe) nessas seções
  const temSecaoLarga = deltaSec > 0 && /<hr[^>]*class="secao"[^>]*data-orient="paisagem"|<hr[^>]*data-orient="paisagem"[^>]*class="secao"/.test(atual.html);
  const cssSecoes = `.ed-texto > [data-secao]:not(table) { margin-right: ${(-deltaSec).toFixed(2)}cm; }
.ed-texto > table[data-secao][style*="width: 100%"], .ed-texto > table[data-secao][style*="width:100%"] { width: calc(100% + ${deltaSec.toFixed(2)}cm) !important; }`;

  useEffect(() => {
    if (!gravar(CHAVE, docs)) setMsg("Atenção: documento grande demais para ficar salvo no navegador (imagens ocupam muito espaço). Salve em Word (.docx) para não perder.");
  }, [docs]);

  // Documento ligado a um arquivo do computador: recupera o vínculo depois de recarregar a página.
  useEffect(() => {
    if (!atual.arquivo || arquivos.current.has(atual.id)) return;
    arquivoLembrado(atual.id).then((h) => { if (h) { arquivos.current.set(atual.id, h); setVersaoArquivos((x) => x + 1); } });
  }, [atual.id, atual.arquivo]);

  // Avisa antes de fechar a aba com alterações não salvas no arquivo.
  useEffect(() => {
    const aviso = (e) => {
      if (docs.some((x) => x.arquivo && x.alterado)) { e.preventDefault(); e.returnValue = ""; }
    };
    window.addEventListener("beforeunload", aviso);
    return () => window.removeEventListener("beforeunload", aviso);
  }, [docs]);

  // Troca de documento: carrega o conteúdo (fora do React para não perder o cursor).
  useEffect(() => {
    if (editor.current) {
      editor.current.innerHTML = atual.html;
      prepararLayout();
      renumerarNotas();
      // O navegador normaliza o HTML; guardar a versão normalizada evita "alteração" falsa.
      if (editor.current.innerHTML !== atual.html) mudarDoc({ html: editor.current.innerHTML }, false);
    }
    if (cab.current) cab.current.innerHTML = atual.cabecalho || "";
    if (rod.current) rod.current.innerHTML = atual.rodape || "";
    document.execCommand("styleWithCSS", false, false);
    document.execCommand("defaultParagraphSeparator", false, "p");
    medirPaginas();
  }, [atualId]);

  useEffect(() => {
    const atualizarEstado = () => {
      const s = document.getSelection();
      if (!editor.current?.contains(s?.anchorNode)) return;
      setEstado({
        bold: document.queryCommandState("bold"),
        italic: document.queryCommandState("italic"),
        underline: document.queryCommandState("underline"),
        strikeThrough: document.queryCommandState("strikeThrough"),
        subscript: document.queryCommandState("subscript"),
        superscript: document.queryCommandState("superscript"),
        insertUnorderedList: document.queryCommandState("insertUnorderedList"),
        insertOrderedList: document.queryCommandState("insertOrderedList"),
        bloco: (document.queryCommandValue("formatBlock") || "p").toLowerCase(),
      });
      // Página onde está o cursor
      if (s.rangeCount) {
        const r = s.getRangeAt(0).getBoundingClientRect();
        const topo = editor.current.getBoundingClientRect().top;
        const y = ((r.top || r.bottom) - topo) / zoomRef.current;
        if (r.top || r.bottom) setPagAtual(quebrasRef.current.filter((q) => q <= y).length + 1);
      }
    };
    document.addEventListener("selectionchange", atualizarEstado);
    return () => document.removeEventListener("selectionchange", atualizarEstado);
  }, []);

  // ----- páginas -----

  const quebrasRef = useRef([]);
  const medindo = useRef(false);
  const [faixas, setFaixas] = useState([]); // trechos em seção mais larga (fundo branco extra)
  const secaoRef = useRef({});
  secaoRef.current = { base: orientBase, H: alturaUtil, Houtra: (altOutra - m.sup - m.inf) * CM, larga: deltaSec > 0 };

  // Marca os parágrafos das seções na outra orientação (só para a tela; o .docx e o PDF
  // recalculam a partir das quebras de seção).
  function marcarSecoes() {
    const ed = editor.current;
    if (!ed) return;
    const { base } = secaoRef.current;
    let orient = base;
    for (const f of ed.children) {
      if (f.matches("hr.secao")) { orient = f.dataset.orient === "paisagem" ? "paisagem" : "retrato"; f.removeAttribute("data-secao"); continue; }
      const v = orient === base ? null : orient === "paisagem" ? "larga" : "estreita";
      if (v) { if (f.dataset.secao !== v) f.dataset.secao = v; } else if (f.hasAttribute("data-secao")) f.removeAttribute("data-secao");
    }
  }

  // Tabulações: cada <span class="tab"> vai até a próxima parada (as do parágrafo, em
  // data-tabs, em cm a partir da margem; depois delas, a cada 1,25 cm como no Word).
  function ajustarTabs(raiz = editor.current) {
    const tabs = raiz ? [...raiz.querySelectorAll("span.tab")] : [];
    if (!tabs.length) return;
    const z = zoomRef.current || 1;
    tabs.forEach((t) => { t.style.width = "0px"; });
    for (const t of tabs) {
      const bloco = t.closest("p,h1,h2,h3,h4,h5,h6,li,blockquote,div:not(.ed-texto):not(.ed-area)") || raiz;
      const cel = t.closest("td,th");
      const ref = cel || raiz;
      const pad = cel ? parseFloat(getComputedStyle(cel).paddingLeft) || 0 : 0;
      const r0 = ref.getBoundingClientRect().left;
      const x = (t.getBoundingClientRect().left - r0) / z - pad;
      const paradas = (bloco.dataset?.tabs || "").split(",").map(Number).filter((v) => v > 0).map((v) => v * CM);
      // Recuo deslocado (ex.: "1.<tab>Texto") também é uma parada, como no Word.
      if (parseFloat(getComputedStyle(bloco).textIndent) < 0) paradas.push((bloco.getBoundingClientRect().left - r0) / z - pad);
      paradas.sort((a, b) => a - b);
      let alvo = paradas.find((v) => v > x + 0.5);
      if (alvo === undefined) { const passo = 1.25 * CM; alvo = (Math.floor((x + 0.5) / passo) + 1) * passo; }
      t.style.width = `${(alvo - x).toFixed(1)}px`;
    }
  }

  function prepararLayout() {
    marcarSecoes();
    ajustarTabs();
  }

  function inserirTab() {
    const s = document.getSelection();
    if (!s.rangeCount || !editor.current.contains(s.anchorNode)) return;
    const r = s.getRangeAt(0);
    r.deleteContents();
    const sp = document.createElement("span");
    sp.className = "tab";
    sp.contentEditable = "false";
    sp.textContent = "\t";
    r.insertNode(sp);
    // Depois da tabulação precisa de um nó de texto para o cursor ficar nele.
    let depois = sp.nextSibling;
    if (!depois || depois.nodeType !== 3) { depois = document.createTextNode("\u200b"); sp.after(depois); }
    s.collapse(depois, depois.data.startsWith("\u200b") ? 1 : 0);
    salvarHtml();
  }

  function inserirQuebraSecao(orient) {
    cmd("insertHTML", `<hr class="secao" data-orient="${orient}"><p><br></p>`);
    setMsg(orient === "paisagem" ? "As páginas depois da quebra ficam em paisagem. Para voltar ao retrato, insira outra quebra de seção \"Retrato\"." : "");
  }
  function medirPaginas() {
    if (medindo.current) return;
    medindo.current = true;
    requestAnimationFrame(() => {
      medindo.current = false;
      const ed = editor.current;
      if (!ed) return;
      const sec = secaoRef.current;
      let H = sec.H;
      const lista = [];
      const novasFaixas = [];
      let faixa = null;
      let inicio = 0;
      for (const filho of ed.children) {
        const topo = filho.offsetTop, fim = topo + filho.offsetHeight;
        if (filho.matches("hr.secao")) {
          // Quebra de seção: começa página nova, com a altura da orientação escolhida.
          if (faixa) faixa.fim = topo;
          faixa = null;
          H = filho.dataset.orient === sec.base || (!filho.dataset.orient && sec.base === "retrato") ? sec.H : sec.Houtra;
          if (sec.larga && filho.dataset.orient === "paisagem") novasFaixas.push((faixa = { topo: fim, fim }));
          lista.push(fim); inicio = fim; continue;
        }
        if (faixa) faixa.fim = fim;
        if (filho.tagName === "HR") { lista.push(fim); inicio = fim; continue; }
        while (fim > inicio + H + 1) {
          // Parágrafo curto que não cabe vai inteiro para a próxima página; longo é dividido.
          const quebra = topo > inicio && fim - topo < 48 ? topo : inicio + H;
          lista.push(quebra);
          inicio = quebra;
        }
      }
      setFaixas((a) => (JSON.stringify(a) === JSON.stringify(novasFaixas) ? a : novasFaixas));
      quebrasRef.current = lista;
      atualizarSumarios(lista);
      setQuebras((a) => (a.length === lista.length && a.every((v, i) => Math.abs(v - lista[i]) < 1) ? a : lista));
    });
  }
  const alturaUtilRef = useRef(alturaUtil);
  alturaUtilRef.current = alturaUtil;
  useLayoutEffect(() => { marcarSecoes(); medirPaginas(); }, [alturaUtil, largCm, m.esq, m.dir]);

  // ----- alterações -----

  function mudarDoc(patch, marcarAlterado = true) {
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, ...patch, atualizado: Date.now(), ...(marcarAlterado ? { alterado: true } : {}) } : x)));
  }

  // HTML do editor sem marcas temporárias da tela (células selecionadas).
  const htmlLimpo = () => editor.current.innerHTML.replace(/ data-sel="1"/g, "");

  function salvarHtml() {
    prepararLayout();
    renumerarNotas();
    mudarDoc({ html: htmlLimpo() });
    medirPaginas();
  }

  function cmd(c, v = null) {
    editor.current.focus();
    document.execCommand(c, false, v);
    if (/List$/.test(c)) tirarListasDeParagrafos();
    salvarHtml();
  }

  // O Chrome cria a lista dentro do <p> (<p><ol>…</ol></p>); tira para fora, mantendo o cursor.
  function tirarListasDeParagrafos() {
    const presas = editor.current.querySelectorAll("p > ol, p > ul");
    if (!presas.length) return;
    const s = document.getSelection();
    const cursor = s.rangeCount ? { n: s.anchorNode, o: s.anchorOffset } : null;
    presas.forEach((l) => {
      const p = l.parentElement;
      const depois = p.cloneNode(false);
      while (l.nextSibling) depois.appendChild(l.nextSibling);
      p.after(l);
      if (depois.textContent.trim() || depois.querySelector("img,ol,ul,table")) l.after(depois);
      if (!p.textContent.trim() && !p.querySelector("img")) p.remove();
    });
    if (cursor?.n?.isConnected) { try { s.collapse(cursor.n, cursor.o); } catch { /* cursor fica onde estiver */ } }
  }

  // Tamanho em pontos (o comando do navegador só conhece 7 tamanhos fixos): marca a seleção
  // com o tamanho 7 e troca essas marcas por <span style="font-size: Npt">.
  function aplicarTamanho(pt) {
    editor.current.focus();
    const s = document.getSelection();
    if (!s.rangeCount || s.isCollapsed) return setMsg("Selecione o texto para mudar o tamanho.");
    document.execCommand("styleWithCSS", false, false);
    document.execCommand("fontSize", false, "7");
    editor.current.querySelectorAll('font[size="7"]').forEach((f) => {
      const sp = document.createElement("span");
      sp.style.fontSize = `${pt}pt`;
      if (f.getAttribute("face")) sp.style.fontFamily = f.getAttribute("face");
      if (f.getAttribute("color")) sp.style.color = f.getAttribute("color");
      sp.append(...f.childNodes);
      f.replaceWith(sp);
      // Tamanhos antigos dentro do trecho deixam de valer
      sp.querySelectorAll("[style*='font-size']").forEach((x) => { x.style.fontSize = ""; });
      sp.querySelectorAll("font[size]").forEach((x) => x.removeAttribute("size"));
    });
    setMsg("");
    salvarHtml();
  }

  // ----- pincel de formatação -----

  function pegarFormato() {
    if (pincel) return setPincel(null);
    const s = document.getSelection();
    if (!s.rangeCount || !editor.current.contains(s.anchorNode)) return setMsg("Clique no texto cuja formatação quer copiar.");
    const el = s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode;
    const st = getComputedStyle(el);
    let fundo = "";
    for (let x = el; x && x !== editor.current; x = x.parentElement) {
      const bg = getComputedStyle(x).backgroundColor;
      if (bg && !/rgba\(0, 0, 0, 0\)|transparent/.test(bg)) { fundo = bg; break; }
    }
    const deco = st.textDecorationLine || "";
    setPincel({
      familia: st.fontFamily.split(",")[0].replace(/["']/g, "").trim(), tamanho: Math.round(parseFloat(st.fontSize) * 0.75 * 2) / 2,
      negrito: Number(st.fontWeight) >= 600, italico: st.fontStyle === "italic", sublinhado: /underline/.test(deco), tachado: /line-through/.test(deco),
      cor: st.color, fundo,
    });
    setMsg("Pincel: agora selecione o texto que vai receber a formatação (Esc cancela).");
  }

  function soltarPincel() {
    if (!pincel) return;
    const s = document.getSelection();
    if (!s.rangeCount || s.isCollapsed || !editor.current.contains(s.anchorNode)) return;
    document.execCommand("removeFormat");
    const f = pincel;
    R.envolverIntervalo(s.getRangeAt(0), () => {
      const sp = document.createElement("span");
      sp.style.fontFamily = `'${f.familia}'`;
      sp.style.fontSize = `${f.tamanho}pt`;
      if (f.negrito) sp.style.fontWeight = "bold";
      if (f.italico) sp.style.fontStyle = "italic";
      const deco = [f.sublinhado && "underline", f.tachado && "line-through"].filter(Boolean).join(" ");
      if (deco) sp.style.textDecoration = deco;
      sp.style.color = f.cor;
      if (f.fundo) sp.style.backgroundColor = f.fundo;
      return sp;
    });
    setPincel(null);
    setMsg("");
    salvarHtml();
  }

  // ----- listas e parágrafos -----

  // Formato da numeração da lista onde está o cursor (cria a lista se ainda não houver).
  function formatoNumeracao(tipo) {
    editor.current.focus();
    const s = document.getSelection();
    const naLista = () => {
      const n = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode);
      let ol = n?.closest?.("ol");
      while (ol?.parentElement?.closest?.("ol") && editor.current.contains(ol.parentElement)) ol = ol.parentElement.closest("ol");
      return ol && editor.current.contains(ol) ? ol : null;
    };
    if (!naLista()) { document.execCommand("insertOrderedList"); tirarListasDeParagrafos(); }
    const ol = naLista();
    if (!ol) return;
    if (tipo === "decimal") ol.removeAttribute("data-num"); else ol.dataset.num = tipo;
    salvarHtml();
  }

  // Bordas e sombreamento dos parágrafos selecionados.
  function bordaParagrafo(tipo) {
    editor.current.focus();
    const lista = blocosSelecionados();
    if (!lista.length) return setMsg("Clique no parágrafo (ou selecione o trecho) que deseja ajustar.");
    const linha = "1px solid #000";
    lista.forEach((el) => {
      el.style.border = "";
      el.style.removeProperty("border-image");
      el.style.padding = "";
      if (tipo === "inferior") { el.style.borderBottom = linha; el.style.paddingBottom = "2pt"; }
      if (tipo === "superior") { el.style.borderTop = linha; el.style.paddingTop = "2pt"; }
      if (tipo === "caixa") { el.style.border = linha; el.style.padding = "3pt 5pt"; }
    });
    salvarHtml();
  }

  function sombreamentoParagrafo(cor) {
    editor.current.focus();
    const lista = blocosSelecionados();
    lista.forEach((el) => { el.style.backgroundColor = cor || ""; });
    salvarHtml();
  }

  // Data de hoje por extenso, no formato usado em petições.
  function textoData(formato) {
    const h = new Date();
    const mes = h.toLocaleDateString("pt-BR", { month: "long" });
    const extenso = `${h.getDate()} de ${mes} de ${h.getFullYear()}`;
    return {
      extenso, curta: h.toLocaleDateString("pt-BR"),
      semana: `${h.toLocaleDateString("pt-BR", { weekday: "long" })}, ${extenso}`,
      local: `São Paulo, ${extenso}.`,
    }[formato];
  }

  // A+ / A-: próximo tamanho da lista a partir do tamanho atual do trecho.
  function tamanhoRelativo(dir) {
    const s = document.getSelection();
    if (!s.rangeCount || s.isCollapsed) return setMsg("Selecione o texto para mudar o tamanho.");
    const el = s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode;
    const atualPt = Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2;
    const prox = dir > 0 ? TAMANHOS.find((t) => t > atualPt) : [...TAMANHOS].reverse().find((t) => t < atualPt);
    if (prox) aplicarTamanho(prox);
  }

  // Maiúsculas/minúsculas no trecho selecionado, mantendo a formatação.
  function trocarCaixa(modo) {
    const s = document.getSelection();
    if (!s.rangeCount || s.isCollapsed || !editor.current.contains(s.anchorNode)) return setMsg("Selecione o texto.");
    const r = s.getRangeAt(0);
    const segs = R.textosNoIntervalo(r);
    let inicioFrase = true;
    for (const t of segs) {
      let txt = t.data.toLocaleLowerCase("pt-BR");
      if (modo === "mai") txt = t.data.toLocaleUpperCase("pt-BR");
      else if (modo === "palavras") txt = txt.replace(/(^|[\s(“"'«-])(\p{L})/gu, (_, a, b) => a + b.toLocaleUpperCase("pt-BR"));
      else if (modo === "frase") txt = txt.replace(/\p{L}/gu, (ch) => { const out = inicioFrase ? ch.toLocaleUpperCase("pt-BR") : ch; inicioFrase = false; return out; }).replace(/([.!?]\s+)(\p{L})/gu, (_, a, b) => a + b.toLocaleUpperCase("pt-BR"));
      t.data = txt;
    }
    if (segs.length) {
      const nr = document.createRange();
      nr.setStartBefore(segs[0]);
      nr.setEndAfter(segs[segs.length - 1]);
      s.removeAllRanges();
      s.addRange(nr);
    }
    salvarHtml();
  }

  function inserirSimbolo(ch) {
    editor.current.focus();
    if (atual.controlar && garantirAutor()) { R.inserirComControle(ch, autor || garantirAutor()); salvarHtml(); }
    else cmd("insertText", ch);
  }

  async function colar() {
    try {
      const t = await navigator.clipboard.readText();
      editor.current.focus();
      if (atual.controlar && autor) { R.inserirComControle(t.replace(/\r?\n/g, " "), autor); salvarHtml(); }
      else cmd("insertText", t);
    } catch {
      setMsg("O navegador não deixou ler a área de transferência: use Ctrl+V.");
    }
  }

  function larguraDaPagina() {
    const disp = (areaRef.current?.clientWidth || 900) - 40;
    const pag = pagRef.current?.getBoundingClientRect().width / zoomRef.current || 800;
    setZoom(Math.max(50, Math.min(200, Math.floor((disp / pag) * 10) * 10)));
  }

  // Parágrafos tocados pela seleção (para espaçamento e recuo).
  function blocosSelecionados() {
    const s = document.getSelection();
    if (!s.rangeCount || !editor.current.contains(s.anchorNode)) return [];
    const r = s.getRangeAt(0);
    const folhas = () => [...editor.current.querySelectorAll("p,h1,h2,h3,h4,li,blockquote,div")].filter((el) => !el.querySelector("p,h1,h2,h3,h4,li,blockquote,div,table") && r.intersectsNode(el));
    let lista = folhas();
    if (!lista.length) { document.execCommand("formatBlock", false, "p"); lista = folhas(); }
    return lista;
  }

  function estiloParagrafo(prop, valor) {
    editor.current.focus();
    const lista = blocosSelecionados();
    if (!lista.length) return setMsg("Clique no parágrafo (ou selecione o trecho) que deseja ajustar.");
    lista.forEach((el) => { el.style[prop] = valor; });
    setMsg("");
    salvarHtml();
  }

  function renomear(nome) {
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, nome } : x)));
  }

  function criar(nome, extra) {
    const n = novoDoc(nome, extra);
    setDocs((ds) => [n, ...ds]);
    setAtualId(n.id);
    return n;
  }

  function trocar(id) {
    if (editor.current.innerHTML !== atual.html) salvarHtml();
    setAtualId(id);
  }

  function excluir() {
    if (!confirm(`Excluir "${atual.nome}"?`)) return;
    arquivos.current.delete(atualId);
    lembrarArquivo(atualId, null);
    const resto = docs.filter((x) => x.id !== atualId);
    const lista = resto.length ? resto : [novoDoc()];
    setDocs(lista);
    setAtualId(lista[0].id);
  }

  // ----- revisão: autor, controle de alterações, comentários -----

  function garantirAutor() {
    if (autor) return autor;
    const nome = prompt("Seu nome (aparece nos comentários e nas alterações):", "")?.trim();
    if (!nome) return null;
    gravar("editorTexto.autor", nome);
    setAutor(nome);
    return nome;
  }

  // O ouvinte é registrado uma vez; por isso lê o estado atual por esta referência.
  const refsRev = useRef({});
  refsRev.current = { controlar: !!atual.controlar, autor, salvar: salvarHtml, desfazer: desfazerControle };
  useEffect(() => {
    const ed = editor.current;
    const antes = (e) => {
      const { controlar, autor: nome } = refsRev.current;
      if (!controlar || !nome) return;
      if (e.inputType === "historyUndo") { e.preventDefault(); return refsRev.current.desfazer(); }
      const foto = ed.innerHTML;
      if (R.tratarEntrada(e, nome)) {
        desfazerRev.current.push(foto);
        if (desfazerRev.current.length > 200) desfazerRev.current.shift();
        refsRev.current.salvar();
      }
    };
    ed.addEventListener("beforeinput", antes);
    return () => ed.removeEventListener("beforeinput", antes);
  }, []);

  function desfazerControle() {
    const foto = desfazerRev.current.pop();
    if (foto === undefined) return;
    editor.current.innerHTML = foto;
    salvarHtml();
  }

  function alternarControle() {
    if (!atual.controlar && !garantirAutor()) return;
    mudarDoc({ controlar: !atual.controlar }, false);
    desfazerRev.current = [];
    setMsg(!atual.controlar ? "Controle de alterações ligado: o que você incluir ou excluir fica marcado com seu nome." : "Controle de alterações desligado.");
  }

  function revisar(el, aceitarIt) {
    if (!el?.isConnected) return;
    if (aceitarIt) R.aceitar(el); else R.rejeitar(el);
    salvarHtml();
  }

  function revisarTodas(aceitarIt) {
    const lista = [...editor.current.querySelectorAll("ins.rev, del.rev")];
    if (!lista.length) return setMsg("Não há alterações marcadas.");
    if (!confirm(`${aceitarIt ? "Aceitar" : "Rejeitar"} todas as ${lista.length} alterações?`)) return;
    lista.forEach((el) => (aceitarIt ? R.aceitar(el) : R.rejeitar(el)));
    salvarHtml();
  }

  function comentar() {
    const s = document.getSelection();
    if (!s.rangeCount || s.isCollapsed || !editor.current.contains(s.anchorNode)) return setMsg("Selecione o trecho que deseja comentar.");
    const nome = garantirAutor();
    if (!nome) return;
    const id = "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
    R.envolverIntervalo(s.getRangeAt(0), () => { const sp = document.createElement("span"); sp.className = "coment"; sp.dataset.c = id; return sp; });
    s.removeAllRanges();
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, comentarios: { ...(x.comentarios || {}), [id]: { autor: nome, data: new Date().toISOString(), texto: "" } }, alterado: true } : x)));
    setComNovo(id);
    setComAtivo(id);
    salvarHtml();
  }

  function mudarComentario(id, patch) {
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, comentarios: { ...(x.comentarios || {}), [id]: { ...(x.comentarios || {})[id], ...patch } }, alterado: true } : x)));
  }

  function excluirComentario(id) {
    editor.current.querySelectorAll(`span.coment[data-c="${id}"]`).forEach(R.desembrulhar);
    setDocs((ds) => ds.map((x) => {
      if (x.id !== atualId) return x;
      const c = { ...(x.comentarios || {}) };
      delete c[id];
      return { ...x, comentarios: c, alterado: true };
    }));
    salvarHtml();
  }

  // ----- notas de rodapé -----

  function renumerarNotas() {
    editor.current?.querySelectorAll("sup.nota").forEach((s, i) => { if (s.textContent !== String(i + 1)) s.textContent = i + 1; });
  }

  function inserirNota() {
    const id = "n" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
    const s = document.getSelection();
    if (!s.rangeCount || !editor.current.contains(s.anchorNode)) {
      // Sem cursor no texto: põe a nota no fim do documento.
      const r = document.createRange();
      r.selectNodeContents(editor.current.lastElementChild || editor.current);
      r.collapse(false);
      s.removeAllRanges();
      s.addRange(r);
    }
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, notas: { ...(x.notas || {}), [id]: "<p><br></p>" }, alterado: true } : x)));
    // Inserção direta no ponto do cursor (o insertHTML do Chrome joga elementos não editáveis
    // para fora do parágrafo quando o cursor está no fim da linha).
    const r = s.getRangeAt(0);
    r.deleteContents();
    const sup = document.createElement("sup");
    sup.className = "nota";
    sup.dataset.nota = id;
    sup.setAttribute("contenteditable", "false");
    sup.textContent = "0";
    const depois = document.createTextNode("\u200b"); // permite pôr o cursor logo após a chamada
    r.insertNode(depois);
    r.insertNode(sup);
    // Desenha o campo da nota já, para o cursor ir para ele antes da próxima tecla.
    flushSync(() => salvarHtml());
    const campo = document.querySelector(`[data-nota-texto="${id}"]`);
    if (campo) {
      campo.focus();
      const r = document.createRange();
      r.selectNodeContents(campo);
      r.collapse(false);
      document.getSelection().removeAllRanges();
      document.getSelection().addRange(r);
    }
  }

  function mudarNota(id, html) {
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, notas: { ...(x.notas || {}), [id]: html }, alterado: true } : x)));
  }

  // ----- sumário -----

  function atualizarSumarios(quebrasAtuais = quebrasRef.current) {
    const ed = editor.current;
    const sumarios = ed?.querySelectorAll("[data-sumario]");
    if (!sumarios?.length) return;
    const titulos = [...ed.querySelectorAll("h1,h2,h3")].filter((h) => !h.closest("[data-sumario]") && h.textContent.trim());
    const itens = titulos.map((h, i) => {
      const pag = quebrasAtuais.filter((q) => q <= h.offsetTop).length + 1;
      return `<div class="sum-item sum-n${h.tagName[1]}" data-alvo="${i}"><span>${esc(h.textContent.trim())}</span><span class="pontos"></span><span>${pag}</span></div>`;
    }).join("") || '<div class="sum-item"><i>Use os estilos Título 1, 2 e 3 para montar o sumário.</i></div>';
    const html = `<div class="sum-titulo">Sumário</div>${itens}`;
    sumarios.forEach((s) => { if (s.innerHTML !== html) s.innerHTML = html; });
  }

  function inserirSumario() {
    cmd("insertHTML", '<div class="sumario" data-sumario="1" contenteditable="false"></div><p><br></p>');
    atualizarSumarios();
    salvarHtml();
  }

  function cliqueEditor(e) {
    const item = e.target.closest?.("[data-alvo]");
    if (item) {
      const titulos = [...editor.current.querySelectorAll("h1,h2,h3")].filter((h) => !h.closest("[data-sumario]") && h.textContent.trim());
      titulos[Number(item.dataset.alvo)]?.scrollIntoView({ block: "center", behavior: "smooth" });
      return;
    }
    const nota = e.target.closest?.("sup.nota");
    if (nota) document.querySelector(`[data-nota-texto="${nota.dataset.nota}"]`)?.focus();
  }

  // ----- imagens -----

  function selecionarImagem(img) {
    // Sem seleção de texto (o navegador pintaria a imagem de azul); Delete é tratado à parte.
    document.getSelection().removeAllRanges();
    setImgSel(img);
  }

  // A moldura acompanha a imagem quando o texto muda ou a janela muda de tamanho.
  useEffect(() => {
    if (!imgSel) return;
    const atualizar = () => { if (!imgSel.isConnected) setImgSel(null); else setTick((x) => x + 1); };
    const fora = (e) => { if (e.target !== imgSel && !e.target.closest?.("[data-moldura]")) setImgSel(null); };
    window.addEventListener("resize", atualizar);
    document.addEventListener("selectionchange", atualizar);
    window.addEventListener("mousedown", fora);
    return () => { window.removeEventListener("resize", atualizar); document.removeEventListener("selectionchange", atualizar); window.removeEventListener("mousedown", fora); };
  }, [imgSel]);

  function redimensionarImagem(e, canto) {
    e.preventDefault();
    e.stopPropagation();
    const img = imgSel;
    const z = zoomRef.current;
    const w0 = img.getBoundingClientRect().width / z, h0 = img.getBoundingClientRect().height / z;
    const proporcao = h0 / w0;
    const max = editor.current.getBoundingClientRect().width / z;
    const x0 = e.clientX;
    const sinal = canto.includes("d") ? 1 : -1; // alças da direita crescem para a direita
    const mover = (ev) => {
      const w = Math.max(24, Math.min(max, w0 + (sinal * (ev.clientX - x0)) / z));
      img.setAttribute("width", Math.round(w));
      img.setAttribute("height", Math.round(w * proporcao));
      img.style.width = img.style.height = "";
      setTick((x) => x + 1);
    };
    const fim = () => { window.removeEventListener("mousemove", mover); window.removeEventListener("mouseup", fim); salvarHtml(); };
    window.addEventListener("mousemove", mover);
    window.addEventListener("mouseup", fim);
  }

  function posicionarImagem(modo) {
    const img = imgSel;
    if (!img) return;
    const bloco = img.closest("p,div,h1,h2,h3,li,td");
    img.style.float = modo === "esq" ? "left" : modo === "dir" ? "right" : "";
    img.style.margin = modo === "esq" ? "0 12px 6px 0" : modo === "dir" ? "0 0 6px 12px" : "";
    if (bloco && bloco !== editor.current) bloco.style.textAlign = modo === "centro" ? "center" : bloco.style.textAlign === "center" ? "" : bloco.style.textAlign;
    salvarHtml();
    setTick((x) => x + 1);
  }

  function tamanhoOriginal() {
    const img = imgSel;
    const max = editor.current.getBoundingClientRect().width / zoomRef.current;
    const w = Math.min(img.naturalWidth, max);
    img.setAttribute("width", Math.round(w));
    img.setAttribute("height", Math.round((w * img.naturalHeight) / img.naturalWidth));
    salvarHtml();
    setTick((x) => x + 1);
  }

  function excluirImagem() {
    imgSel?.remove();
    setImgSel(null);
    salvarHtml();
  }

  // ----- tabelas -----

  function inserirTabela(l, c) {
    setSeletorTabela(null);
    if (!(l > 0 && c > 0)) return;
    cmd("insertHTML", T.novaTabela(l, c));
  }

  function limparSelCelulas() {
    editor.current?.querySelectorAll("[data-sel]").forEach((x) => x.removeAttribute("data-sel"));
    selCelulas.current = selCelulas.current && { ...selCelulas.current, ret: null };
  }

  const celulaDe = (alvo) => {
    const td = alvo?.closest?.("td,th");
    return td && editor.current.contains(td) ? td : null;
  };

  function mouseDownEditor(e) {
    setMenu(null);
    // Clique em imagem: seleciona (moldura com alças); Delete apaga, como no Word.
    if (e.target.tagName === "IMG") {
      e.preventDefault();
      selecionarImagem(e.target);
      return;
    }
    setImgSel(null);
    const td = celulaDe(e.target);
    if (e.button !== 0) return;
    limparSelCelulas();
    selCelulas.current = null;
    if (!td) return;
    // Perto da borda direita da célula: arrasta para mudar a largura da coluna.
    const caixa = td.getBoundingClientRect();
    if (e.clientX > caixa.right - 5) {
      const tabela = td.closest("table");
      const g = T.grade(tabela);
      const c = T.posicao(g, td).c + td.colSpan - 1;
      if (T.iniciarRedimensionar(tabela, c, e.clientX, salvarHtml)) { e.preventDefault(); return; }
    }
    selCelulas.current = { tabela: td.closest("table"), inicio: td, ret: null };
  }

  function mouseMoveEditor(e) {
    const td = celulaDe(e.target);
    if (!e.buttons) {
      const borda = td && e.clientX > td.getBoundingClientRect().right - 5 && td.closest("table") && T.posicao(T.grade(td.closest("table")), td).c + td.colSpan < T.numColunas(T.grade(td.closest("table")));
      editor.current.style.cursor = borda ? "col-resize" : "";
      return;
    }
    const s = selCelulas.current;
    if (!s || !td || td.closest("table") !== s.tabela || td === s.inicio && !s.ret) return;
    // Arrastando de uma célula para outra: seleciona o retângulo de células.
    const ret = T.retangulo(s.tabela, s.inicio, td);
    editor.current.querySelectorAll("[data-sel]").forEach((x) => x.removeAttribute("data-sel"));
    ret.celulas.forEach((x) => x.setAttribute("data-sel", "1"));
    s.ret = ret;
    document.getSelection().removeAllRanges();
    e.preventDefault();
  }

  function menuContexto(e) {
    const td = celulaDe(e.target);
    if (!td) return; // fora de tabela: menu normal do navegador (corretor ortográfico)
    e.preventDefault();
    if (!td.hasAttribute("data-sel")) limparSelCelulas();
    setMenu({ x: Math.min(e.clientX, window.innerWidth - 250), y: Math.min(e.clientY, window.innerHeight - 420), td });
  }

  function opTabela(fn) {
    const td = menu?.td;
    setMenu(null);
    if (!td?.isConnected) return;
    const tabela = td.closest("table");
    fn(tabela, td, selCelulas.current?.ret);
    limparSelCelulas();
    salvarHtml();
  }

  function teclaTabela(e) {
    if (pincel && e.key === "Escape") { setPincel(null); setMsg(""); return; }
    if (imgSel && (e.key === "Delete" || e.key === "Backspace")) { e.preventDefault(); return excluirImagem(); }
    {
      // Tab / Shift+Tab num item de lista (fora de tabela): desce / sobe um nível.
      const s0 = document.getSelection();
      const n0 = s0.rangeCount && (s0.anchorNode?.nodeType === 3 ? s0.anchorNode.parentElement : s0.anchorNode);
      if (e.key === "Tab" && n0?.closest?.("li") && !n0.closest("td,th")) {
        e.preventDefault();
        document.execCommand(e.shiftKey ? "outdent" : "indent");
        return salvarHtml();
      }
      // Tab no texto: tabulação até a próxima parada da régua.
      if (e.key === "Tab" && !e.shiftKey && !e.ctrlKey && !e.altKey && n0 && editor.current.contains(n0) && !n0.closest("td,th")) {
        e.preventDefault();
        return inserirTab();
      }
    }
    if (imgSel && e.key === "Escape") return setImgSel(null);
    const s = document.getSelection();
    const td = s.rangeCount && celulaDe(s.anchorNode?.nodeType === 1 ? s.anchorNode : s.anchorNode?.parentElement);
    // Delete com células selecionadas: apaga o conteúdo delas, como no Word.
    if ((e.key === "Delete" || e.key === "Backspace") && selCelulas.current?.ret) {
      e.preventDefault();
      selCelulas.current.ret.celulas.forEach((x) => { x.innerHTML = "<br>"; });
      limparSelCelulas();
      return salvarHtml();
    }
    if (e.key !== "Tab" || !td) return;
    // Tab / Shift+Tab: próxima / anterior célula; Tab na última cria linha nova.
    e.preventDefault();
    const tabela = td.closest("table");
    const todas = [...tabela.querySelectorAll("td,th")];
    let alvo = todas[todas.indexOf(td) + (e.shiftKey ? -1 : 1)];
    if (!alvo && !e.shiftKey) { alvo = T.inserirLinha(tabela, todas[todas.length - 1], true); salvarHtml(); }
    if (!alvo) return;
    const r = document.createRange();
    r.selectNodeContents(alvo);
    r.collapse(false);
    s.removeAllRanges();
    s.addRange(r);
  }

  function inserirLink() {
    const url = prompt("Endereço do link:", "https://");
    if (url && /^(https?:|mailto:)/i.test(url)) cmd("createLink", url);
  }

  function inserirImagem(e) {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    if (!/^image\/(png|jpe?g|gif|bmp)$/.test(f.type)) return setMsg("Use imagens PNG, JPG, GIF ou BMP.");
    const leitor = new FileReader();
    leitor.onload = () => {
      const img = new Image();
      img.onload = () => {
        const max = (largCm - m.esq - m.dir) * CM;
        const w = Math.min(img.naturalWidth, max);
        const h = Math.round((img.naturalHeight * w) / img.naturalWidth);
        cmd("insertHTML", `<img src="${leitor.result}" width="${Math.round(w)}" height="${h}">`);
      };
      img.src = leitor.result;
    };
    leitor.readAsDataURL(f);
  }

  // Abrir: no Chrome/Edge o documento fica ligado ao arquivo (Salvar grava nele).
  async function abrir() {
    if (!suportaArquivoLocal) return arquivo.current.click();
    try {
      const r = await escolherArquivo([TIPOS.docx, TIPOS.texto]);
      if (!r) return;
      for (const [id, h] of arquivos.current) {
        if (docs.some((x) => x.id === id) && (await h.isSameEntry?.(r.handle))) {
          trocar(id);
          return setMsg(`"${r.handle.name}" já está aberto.`);
        }
      }
      await carregarArquivo(r.file, r.handle);
    } catch (err) {
      setMsg(err.message || "Não foi possível abrir o arquivo.");
    }
  }

  async function abrirArquivo(e) {
    const f = e.target.files[0];
    e.target.value = "";
    if (f) await carregarArquivo(f, null);
  }

  async function carregarArquivo(f, handle) {
    const nome = f.name.replace(/\.[^.]+$/, "");
    const vinculo = handle ? { arquivo: f.name, alterado: false } : {};
    let novo = null;
    try {
      setMsg("Abrindo arquivo...");
      if (/\.docx$/i.test(f.name)) {
        const buf = await f.arrayBuffer();
        let conv;
        try {
          conv = await docxParaHtml(buf);
        } catch {
          // Plano B: leitura simples (só texto, negrito, listas e tabelas).
          const mammoth = await import("mammoth");
          conv = { html: (await mammoth.convertToHtml({ arrayBuffer: buf })).value };
        }
        const avisos = handle ? await recursosNaoSuportados(buf, "docx") : [];
        novo = criar(nome, { html: sanitizar(conv.html || "<p><br></p>"), cabecalho: sanitizar(conv.cabecalho || ""), rodape: sanitizar(conv.rodape || ""), numPag: conv.numPag || "", primeiraDiferente: !!conv.primeiraDiferente, cabecalho1: sanitizar(conv.cabecalho1 || ""), rodape1: sanitizar(conv.rodape1 || ""), pagina: conv.pagina || PAGINA_PADRAO, comentarios: conv.comentarios || {}, notas: Object.fromEntries(Object.entries(conv.notas || {}).map(([k, v]) => [k, sanitizar(v)])), ...vinculo, avisos });
      } else if (/\.(txt|md)$/i.test(f.name)) {
        const t = await f.text();
        novo = criar(nome, { html: t.split(/\r?\n/).map((l) => `<p>${esc(l) || "<br>"}</p>`).join(""), ...vinculo });
      } else if (/\.html?$/i.test(f.name)) {
        novo = criar(nome, { html: sanitizar(await f.text()), ...vinculo });
      } else {
        throw new Error("Formato não suportado. Use .docx, .txt ou .html (arquivos .doc antigos: salve como .docx no Word antes).");
      }
      if (handle && novo) {
        arquivos.current.set(novo.id, handle);
        lembrarArquivo(novo.id, handle);
      }
      setMsg(novo?.avisos?.length ? `Atenção: este arquivo tem ${novo.avisos.join(", ")}, que o editor ainda não mantém. Ao salvar por cima eles seriam perdidos; você será avisado antes.` : "");
    } catch (err) {
      setMsg(err.message || "Não foi possível abrir o arquivo.");
    }
  }

  // Salvar: grava por cima do arquivo aberto (Chrome/Edge). Sem arquivo ligado, pergunta onde
  // salvar; em navegadores sem esse recurso, baixa um .docx.
  async function salvarDocx(comoNovo = false) {
    // Imagens precisam de tamanho explícito no .docx
    editor.current.querySelectorAll("img").forEach((img) => {
      if (!img.getAttribute("width")) img.setAttribute("width", img.clientWidth || img.naturalWidth || 300);
      if (!img.getAttribute("height")) img.setAttribute("height", img.clientHeight || img.naturalHeight || 200);
    });
    renumerarNotas();
    const conteudo = { html: htmlLimpo(), cabecalho: cab.current.innerHTML, rodape: rod.current.innerHTML, cabecalho1: atual.cabecalho1 || "", rodape1: atual.rodape1 || "", notas: atual.notas || {}, comentarios: atual.comentarios || {} };
    const doc = atual;
    const nomeBase = (doc.nome || "documento").replace(/\.docx$/i, "");
    try {
      let handle = comoNovo ? null : arquivos.current.get(doc.id);
      if (handle && !/\.docx$/i.test(handle.name)) handle = null; // .txt/.html: salva como .docx
      if (handle && doc.avisos?.length && !doc.avisosAceitos) {
        const ok = confirm(`"${handle.name}" tem ${doc.avisos.join(", ")}, que este editor não mantém. Se salvar por cima, isso será perdido no arquivo.\n\nOK = salvar por cima mesmo assim\nCancelar = "Salvar como…" outro arquivo (o original fica intacto)`);
        if (!ok) return salvarDocx(true);
      }
      if (!handle && !comoNovo && doc.arquivo && !arquivos.current.has(doc.id) && suportaArquivoLocal) {
        setMsg(`Não encontrei mais o vínculo com "${doc.arquivo}". Escolha onde salvar.`);
      }
      setMsg("Salvando...");
      const blob = await htmlParaDocx(doc, conteudo);
      if (handle) {
        await gravarNoArquivo(handle, blob);
      } else if (suportaArquivoLocal) {
        handle = await salvarComo(blob, `${nomeBase}.docx`, [TIPOS.docx]);
        if (!handle) return setMsg("");
        arquivos.current.set(doc.id, handle);
        lembrarArquivo(doc.id, handle);
      } else {
        baixar(blob, `${nomeBase}.docx`);
      }
      const hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
      setDocs((ds) => ds.map((x) => (x.id === doc.id ? {
        ...x, ...conteudo, alterado: false,
        ...(handle ? { arquivo: handle.name, nome: handle.name.replace(/\.docx$/i, ""), avisos: comoNovo || !x.arquivo || x.arquivo !== handle.name ? [] : x.avisos, avisosAceitos: true } : {}),
      } : x)));
      setMsg(handle ? `Salvo em "${handle.name}" às ${hora}.` : "Arquivo .docx baixado (este navegador não permite salvar no mesmo arquivo; use Chrome ou Edge no computador).");
    } catch (err) {
      setMsg("Erro ao salvar: " + err.message);
    }
  }

  // Impressão e PDF: o paged.js monta as páginas de verdade (como o Word) na janela de
  // visualização: cabeçalho/rodapé em cada página (1ª página diferente), "Página X de Y",
  // notas no pé da própria página e sumário com os números de página exatos.
  // Página de impressão (HTML completo, sem scripts) usada pela janela de impressão e pelo
  // PDF gerado no servidor. `fontes` é o CSS das fontes (com endereços ou já embutidas).
  function montarImpressao(fontes = CSS_FONTES.replace(/url\("\//g, `url("${location.origin}/`)) {
    const fonteN = `'${estilos.normal.fonte}', sans-serif`;
    // Corpo: notas viram notas de rodapé do paged.js; sumário aponta para os títulos.
    const corpo = document.createElement("div");
    corpo.innerHTML = htmlLimpo().replace(/\u200b/g, "");
    corpo.querySelectorAll("sup.nota[data-nota]").forEach((sup) => {
      const n = document.createElement("span");
      n.className = "nota-pdf";
      n.innerHTML = (atual.notas?.[sup.dataset.nota] || "").replace(/<\/p>\s*<p[^>]*>/g, "<br>").replace(/<\/?p[^>]*>/g, "");
      sup.replaceWith(n);
    });
    const titulos = [...corpo.querySelectorAll("h1,h2,h3")].filter((h) => !h.closest("[data-sumario]") && h.textContent.trim());
    titulos.forEach((h, i) => { h.id = `titulo-${i}`; });
    corpo.querySelectorAll("[data-sumario] [data-alvo]").forEach((item) => {
      const num = item.lastElementChild;
      if (num) num.outerHTML = `<a class="sum-pag" href="#titulo-${item.dataset.alvo}"></a>`;
    });
    corpo.querySelectorAll("[contenteditable]").forEach((x) => x.removeAttribute("contenteditable"));
    // Marcações de revisão: como no Word, saem no PDF quando estão visíveis na tela
    // (Exibir › Marcações de revisão). Comentários viram [C1], [C2]… com a lista no fim.
    const comMarcas = mostrarMarcas && (!!corpo.querySelector("ins.rev, del.rev") || corpo.querySelector("span.coment") !== null);
    let listaComentarios = "";
    if (comMarcas) {
      const ordem = [...new Set([...corpo.querySelectorAll("span.coment[data-c]")].map((x) => x.dataset.c))].filter((id) => comentarios[id] && !comentarios[id].resolvido);
      ordem.forEach((id, i) => {
        const trechos = corpo.querySelectorAll(`span.coment[data-c="${CSS.escape(id)}"]`);
        trechos[trechos.length - 1].insertAdjacentHTML("afterend", `<sup class="com-ref">[C${i + 1}]</sup>`);
      });
      const quando = (d) => (d ? new Date(d).toLocaleDateString("pt-BR") : "");
      if (ordem.length) listaComentarios = `<div class="lista-com"><p class="lista-com-titulo">Comentários</p>${ordem.map((id, i) => {
        const c = comentarios[id];
        const resp = (c.respostas || []).map((r) => `<br>↳ <b>${esc(r.autor || "")}</b> (${quando(r.data)}): ${esc(r.texto || "")}`).join("");
        return `<p><b>[C${i + 1}] ${esc(c.autor || "")}</b> (${quando(c.data)}): ${esc(c.texto || "")}${resp}</p>`;
      }).join("")}</div>`;
    } else corpo.querySelectorAll("span.coment").forEach((x) => x.classList.remove("coment"));
    // Fica dentro do corpo (e da última seção), no fim do documento.
    if (listaComentarios) corpo.insertAdjacentHTML("beforeend", listaComentarios);
    const numero = atual.numPag
      ? `<div style="text-align:${atual.numPag === "direita" ? "right" : "center"};font:10pt ${fonteN}">${atual.numPag === "xdey" ? 'Página <span class="pg"></span> de <span class="pgs"></span>' : '<span class="pg"></span>'}</div>`
      : "";
    const area = (cls, html) => `<div class="${cls}"><div class="ed-area">${html || ""}</div></div>`;
    const dif = !!atual.primeiraDiferente;
    const areas = area("cab", cab.current.innerHTML) + area("rod", rod.current.innerHTML + numero) + (dif ? area("cab1", atual.cabecalho1) + area("rod1", atual.rodape1) : "");
    const comSecoes = !!corpo.querySelector(":scope > hr.secao");
    // Quebras de seção: cada trecho numa caixa; os da outra orientação usam a página "outra".
    if (comSecoes) {
      const secoes = [{ orient: orientBase, nos: [] }];
      for (const n of [...corpo.childNodes]) {
        if (n.nodeType === 1 && n.matches("hr.secao")) secoes.push({ orient: n.dataset.orient === "paisagem" ? "paisagem" : "retrato", nos: [] });
        else secoes[secoes.length - 1].nos.push(n);
      }
      corpo.innerHTML = "";
      secoes.forEach((sc, i) => {
        const caixa = document.createElement("div");
        caixa.className = sc.orient === orientBase ? "sec-base" : "sec-outra";
        // Mudança de orientação já começa página nova; mesma orientação precisa da quebra.
        if (i > 0 && sc.orient === secoes[i - 1].orient) caixa.style.breakBefore = "page";
        caixa.append(...sc.nos);
        corpo.appendChild(caixa);
      });
      // Cabeçalho/rodapé dentro da 1ª caixa: fora dela a 1ª página sairia sem nome (e em branco).
      corpo.firstElementChild.insertAdjacentHTML("afterbegin", areas);
    }
    // O paged.js desenha as páginas da outra orientação, mas imprime tudo num só tamanho:
    // depois de montadas, as páginas "outra" ganham o próprio @page.
    const cssOutra = `<style>@page outra { size: ${largOutra}cm ${altOutra}cm; margin: 0; }
        .pagedjs_page.pagedjs_outra_page { page: outra; width: ${largOutra}cm; height: ${altOutra}cm;
          --pagedjs-width: ${largOutra}cm; --pagedjs-height: ${altOutra}cm; --pagedjs-width-right: ${largOutra}cm; --pagedjs-height-right: ${altOutra}cm;
          --pagedjs-width-left: ${largOutra}cm; --pagedjs-height-left: ${altOutra}cm; }</style>`;
    const html = `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>${esc(atual.nome)}</title>
      <style>
        @page { size: ${largCm}cm ${altCm}cm; margin: ${m.sup}cm ${m.dir}cm ${m.inf}cm ${m.esq}cm;
          @top-center { content: element(cabecalho); vertical-align: bottom; padding-bottom: 8pt; }
          @bottom-center { content: element(rodape); vertical-align: top; padding-top: 8pt; }
          @footnote { border-top: 0.75pt solid #000; padding-top: 3pt; margin-top: 8pt; max-width: 5cm; } }
        @page outra { size: ${largOutra}cm ${altOutra}cm; } @page base { size: ${largCm}cm ${altCm}cm; }
        .sec-outra { page: outra; } .sec-base { page: base; }
        ${dif ? "@page :first { @top-center { content: element(cabecalho1); } @bottom-center { content: element(rodape1); } }" : ""}
        ${fontes}
        body { margin: 0; } ${CSS_PAGINA} ${CSS_LISTAS} ${cssDoc}
        .cab { position: running(cabecalho); } .rod { position: running(rodape); }
        .cab1 { position: running(cabecalho1); } .rod1 { position: running(rodape1); }
        .cab, .rod, .cab1, .rod1 { font-family: ${fonteN}; font-size: ${estilos.normal.tamanho}pt; width: 100%; text-align: left; }
        .ed-area { border: 0; padding: 0; } .ed-area p { margin: 0; text-indent: 0; }
        .pg::after { content: counter(page); } .pgs::after { content: counter(pages); }
        .nota-pdf { float: footnote; font: 10pt/1.2 ${fonteN}; text-align: justify; text-indent: 0; }
        .nota-pdf::footnote-call { content: counter(footnote); vertical-align: super; font-size: 0.65em; line-height: 0; }
        .nota-pdf::footnote-marker { content: counter(footnote) " "; vertical-align: super; font-size: 0.7em; }
        .sumario { border: 0; padding: 0; } .sumario a.sum-pag { color: inherit; text-decoration: none; }
        .sumario a.sum-pag::after { content: target-counter(attr(href), page); }
        .ed-texto hr { break-after: page; border: 0; margin: 0; }
        ${comMarcas ? `.ed-texto del.rev { display: inline; text-decoration: line-through; } .ed-texto ins.rev { text-decoration: underline; }
          ${R.cssAutores(autoresRev)}
          .ed-texto span.coment { background: #fff1c2; } .com-ref { font-size: 7pt; color: #9a6e1e; vertical-align: super; line-height: 0; }
          .lista-com { margin-top: 18pt; border-top: 0.75pt solid #999; padding-top: 4pt; font: 9pt/1.3 ${fonteN}; }
          .lista-com p { margin: 0 0 4pt; text-indent: 0; text-align: left; } .lista-com .lista-com-titulo { font-weight: bold; }` : ""}
      </style></head><body>
      ${comSecoes ? "" : areas}
      <div class="ed-texto">${corpo.innerHTML}</div>
      </body></html>`;
    return { html, comMarcas, cssOutra };
  }

  // Janela de impressão (Ctrl+P): o paged.js monta as páginas e abre o diálogo do navegador.
  function imprimir(pdf = false) {
    const w = window.open("", "_blank");
    if (!w) return setMsg("Permita pop-ups para imprimir ou salvar em PDF.");
    const { html, comMarcas, cssOutra } = montarImpressao();
    w.document.write(html);
    w.document.close();
    const barra = `<div style="position:fixed;top:0;left:0;right:0;z-index:10;display:flex;align-items:center;gap:12px;padding:10px 16px;background:#8a6a2c;color:#fff;font:14px sans-serif;box-shadow:0 2px 6px rgba(0,0,0,0.3)">
      <b>${esc(atual.nome)}</b><span id="estado-pdf">Montando as páginas…</span><span style="flex:1"></span>
      <button id="btn-pdf" style="padding:6px 14px;border:0;border-radius:5px;background:#fff;color:#8a6a2c;font-weight:600;cursor:pointer">${pdf ? "Salvar como PDF" : "Imprimir"}</button></div>`;
    // Depois que o paged.js monta as páginas: barra no topo e janela de impressão.
    w.prontoParaImprimir = (total) => {
      w.document.head.insertAdjacentHTML("beforeend", cssOutra);
      // Só na tela: folhas com sombra sobre fundo escuro, como na visualização de impressão do Word.
      w.document.body.insertAdjacentHTML("afterbegin", `<style>@media print { #barra-pdf { display: none !important; } }
        @media screen { html, body { background: #6b7468; } .pagedjs_pages { display: flex; flex-direction: column; align-items: center; gap: 16px; padding: 64px 0 24px; }
        .pagedjs_page { background: #fff; box-shadow: 0 2px 12px rgba(0,0,0,0.35); } }</style><div id="barra-pdf">${barra}</div>`);
      w.document.getElementById("estado-pdf").textContent = `${total} página${total === 1 ? "" : "s"}${comMarcas ? " · com marcações de revisão" : ""}${pdf ? ' — na janela de impressão escolha Destino: "Salvar como PDF"' : ""}`;
      w.document.getElementById("btn-pdf").onclick = () => w.print();
      w.focus();
      if (!w.semDialogo) w.print();
    };
    const sc = w.document.createElement("script");
    sc.textContent = `window.PagedConfig = { auto: false };`;
    w.document.head.appendChild(sc);
    const pg = w.document.createElement("script");
    pg.src = new URL(pagedUrl, location.href).href;
    pg.onload = async () => {
      try {
        await w.document.fonts.ready;
        const fluxo = await w.PagedPolyfill.preview();
        w.prontoParaImprimir(fluxo.total);
      } catch (err) {
        w.document.body.insertAdjacentHTML("afterbegin", `<p style="color:#fff;font:14px sans-serif;padding:16px">Não foi possível montar as páginas (${esc(err.message)}).</p>`);
      }
    };
    pg.onerror = () => w.document.body.insertAdjacentHTML("afterbegin", '<p style="color:#fff;font:14px sans-serif;padding:16px">Não foi possível carregar o montador de páginas.</p>');
    w.document.head.appendChild(pg);
  }

  // Exportar PDF: o servidor gera o arquivo (Chromium) e ele é baixado direto. A página vai
  // completa, com o montador de páginas e as fontes embutidos (o servidor não acessa a rede).
  const [gerandoPdf, setGerandoPdf] = useState(false);
  async function exportarPdf() {
    if (gerandoPdf) return;
    setGerandoPdf(true);
    setMsg("Gerando o PDF…");
    try {
      const [polyfill, fontes] = await Promise.all([recursoTexto(pagedUrl), fontesEmbutidas()]);
      const montada = montarImpressao("/*FONTES*/");
      const { cssOutra } = montada;
      const extras = await fontesSubstitutas(montada.html);
      const html = montada.html.replace("/*FONTES*/", () => fontes + extras);
      const scripts = `<script>window.PagedConfig = { auto: false };</script><script>${polyfill.replace(/<\/script/gi, "<\\/script")}</script>
        <script>(async () => { try { await document.fonts.ready; const f = await PagedPolyfill.preview();
          document.head.insertAdjacentHTML("beforeend", ${JSON.stringify(cssOutra)}); window.__paginas = f.total; window.__pdfPronto = true;
        } catch (e) { window.__pdfErro = String((e && e.message) || e); } })();</script>`;
      // Função no replace: o texto do paged.js tem "$&" e afins, que o replace interpretaria.
      const pagina = html.replace("</body></html>", () => `${scripts}</body></html>`);
      // A Vercel aceita até 4,5 MB por envio (documentos com muitas imagens passam disso).
      if (pagina.length > 4.3e6) throw new Error("o documento tem imagens demais para gerar no servidor");
      const r = await fetch("/api/pdf", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ html: pagina }) });
      if (!r.ok || !/pdf/.test(r.headers.get("content-type") || "")) {
        const j = await r.json().catch(() => ({}));
        throw new Error(j.error || (r.status === 404 ? "o servidor de PDF não está disponível aqui" : `erro ${r.status}`));
      }
      baixar(await r.blob(), `${(atual.nome || "documento").replace(/[\\/:*?"<>|]/g, "-")}.pdf`);
      setMsg("PDF baixado.");
    } catch (e) {
      setMsg(`Não foi possível gerar o PDF direto (${e.message}). Use Arquivo › Imprimir (Ctrl+P) e escolha "Salvar como PDF".`);
    } finally {
      setGerandoPdf(false);
    }
  }

  // ----- localizar e substituir -----

  function textos() {
    const w = document.createTreeWalker(editor.current, NodeFilter.SHOW_TEXT);
    const nos = [];
    let total = "";
    for (let n = w.nextNode(); n; n = w.nextNode()) { nos.push({ n, ini: total.length }); total += n.data; }
    return { nos, total };
  }
  function ocorrencias(termo, caso) {
    if (!termo) return { lista: [], nos: [] };
    const { nos, total } = textos();
    const alvo = caso ? total : total.toLocaleLowerCase("pt-BR");
    const t = caso ? termo : termo.toLocaleLowerCase("pt-BR");
    const lista = [];
    for (let i = alvo.indexOf(t); i >= 0; i = alvo.indexOf(t, i + t.length)) lista.push(i);
    return { lista, nos };
  }
  function intervalo(nos, ini, fim) {
    const ponto = (pos, final) => {
      for (let i = nos.length - 1; i >= 0; i--) {
        if (nos[i].ini < pos || (!final && nos[i].ini === pos)) return { n: nos[i].n, o: pos - nos[i].ini };
      }
      return { n: nos[0].n, o: 0 };
    };
    const a = ponto(ini, false), b = ponto(fim, true);
    const r = document.createRange();
    r.setStart(a.n, a.o);
    r.setEnd(b.n, b.o);
    return r;
  }
  function posicaoCursor(nos) {
    const s = document.getSelection();
    if (!s.rangeCount || !editor.current.contains(s.focusNode)) return 0;
    const r = s.getRangeAt(0);
    const no = nos.find((x) => x.n === r.endContainer);
    return no ? no.ini + r.endOffset : 0;
  }
  function selecionar(r) {
    editor.current.focus();
    const s = document.getSelection();
    s.removeAllRanges();
    s.addRange(r);
    const el = r.startContainer.parentElement;
    el?.scrollIntoView({ block: "center" });
  }
  // A ocorrência fica destacada e o cursor continua na caixa de busca: Enter vai para a
  // próxima (antes o foco ia para o texto e o Enter seguinte apagava a palavra achada).
  function destacar(r) {
    achado.current = r;
    try {
      if (r && window.Highlight && CSS.highlights) CSS.highlights.set("busca", new Highlight(r));
      else CSS.highlights?.delete("busca");
    } catch { /* navegador sem destaque */ }
    r?.startContainer.parentElement?.scrollIntoView({ block: "center" });
  }
  function fecharBusca() {
    const r = achado.current;
    destacar(null);
    setBusca(null);
    // Como no Word: ao fechar, o trecho achado fica selecionado no texto.
    if (r && editor.current.contains(r.startContainer) && !r.collapsed) selecionar(r);
  }
  function localizar(paraTras = false) {
    const { termo, caso } = busca;
    const { lista, nos } = ocorrencias(termo, caso);
    if (!lista.length) { destacar(null); return setBusca((b) => ({ ...b, info: "Nada encontrado" })); }
    const s = document.getSelection();
    let pos = posicaoCursor(nos);
    const a = achado.current;
    if (a && editor.current.contains(a.startContainer)) {
      const no = (x) => nos.find((y) => y.n === x);
      pos = paraTras ? (no(a.startContainer)?.ini ?? 0) + a.startOffset : (no(a.endContainer)?.ini ?? 0) + a.endOffset;
    } else if (paraTras && s.rangeCount) pos -= s.toString().length;
    let i = paraTras ? [...lista].reverse().findIndex((x) => x < pos) : lista.findIndex((x) => x >= pos);
    if (paraTras) i = i < 0 ? lista.length - 1 : lista.length - 1 - i;
    else if (i < 0) i = 0;
    destacar(intervalo(nos, lista[i], lista[i] + termo.length));
    setBusca((b) => ({ ...b, info: `${i + 1} de ${lista.length}` }));
    campoBusca.current?.focus();
  }
  function iguais(a, b, caso) {
    return caso ? a === b : a.toLocaleLowerCase("pt-BR") === b.toLocaleLowerCase("pt-BR");
  }
  function substituir() {
    const { termo, novo, caso } = busca;
    const a = achado.current;
    if (a && editor.current.contains(a.startContainer) && iguais(a.toString(), termo, caso)) {
      selecionar(a);
      document.execCommand("insertText", false, novo);
      achado.current = null;
      salvarHtml();
    }
    localizar();
  }
  function substituirTudo() {
    const { termo, novo, caso } = busca;
    const { lista, nos } = ocorrencias(termo, caso);
    // Do fim para o começo, para as posições anteriores não mudarem.
    for (let i = lista.length - 1; i >= 0; i--) {
      selecionar(intervalo(nos, lista[i], lista[i] + termo.length));
      document.execCommand("insertText", false, novo);
    }
    salvarHtml();
    destacar(null);
    setBusca((b) => ({ ...b, info: lista.length ? `${lista.length} substituição(ões) feitas` : "Nada encontrado" }));
  }
  function abrirBusca(substituirTambem) {
    const s = document.getSelection();
    const sel = s.rangeCount && editor.current.contains(s.anchorNode) ? s.toString() : "";
    setBusca((b) => ({ termo: sel || b?.termo || "", novo: b?.novo || "", caso: b?.caso || false, info: "", subst: substituirTambem || b?.subst }));
    setTimeout(() => campoBusca.current?.focus(), 0);
  }

  function teclas(e) {
    const noTexto = editor.current?.contains(e.target);
    if (noTexto) teclaTabela(e);
    if (!e.defaultPrevented && e.key === "Escape" && busca) { e.preventDefault(); return fecharBusca(); }
    if (e.defaultPrevented || !(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    // Atalhos do Word (inglês) para alinhamento, quebra de página e tamanho da fonte.
    if (noTexto && !e.altKey) {
      const acao = {
        e: () => cmd("justifyCenter"), j: () => cmd("justifyFull"), r: () => cmd("justifyRight"), l: () => cmd("justifyLeft"),
        enter: () => cmd("insertHTML", "<hr><p><br></p>"),
        "]": () => tamanhoRelativo(1), "[": () => tamanhoRelativo(-1), ">": () => tamanhoRelativo(1), "<": () => tamanhoRelativo(-1),
      }[e.shiftKey && (k === "." || k === ",") ? (k === "." ? ">" : "<") : k];
      if (acao && !(e.shiftKey && /^[ejrl]$/.test(k))) { e.preventDefault(); return acao(); }
    }
    if (k === "s") { e.preventDefault(); salvarDocx(e.shiftKey); }
    if (k === "p") { e.preventDefault(); imprimir(); }
    if (k === "f") { e.preventDefault(); abrirBusca(false); }
    if (k === "h") { e.preventDefault(); abrirBusca(true); }
    if (k === "z" && !e.shiftKey && atual.controlar && editor.current?.contains(e.target)) { e.preventDefault(); desfazerControle(); }
  }

  const arquivoAtual = atual.arquivo && (arquivos.current.get(atual.id)?.name || atual.arquivo);
  const palavras = (atual.html.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").match(/\S+/g) || []).length;
  const totalPag = quebras.length + 1;
  const textoNumero = { direita: "1", centro: "1", xdey: `Página 1 de ${totalPag}` }[atual.numPag];

  return (
    <div style={{ fontFamily: FONTE_UI, borderRadius: 10, boxShadow: "0 2px 14px rgba(0,0,0,0.18)" }} onKeyDown={teclas}>
      <style>{CSS_FONTES + CSS_PAGINA + CSS_LISTAS + cssDoc + CSS_TELA + cssSecoes + cssRevisao + (mostrarGrade ? "" : ".ed-texto td, .ed-texto th { outline: none; }")}</style>
      <input ref={arquivo} type="file" accept=".docx,.txt,.md,.html,.htm" onChange={abrirArquivo} style={{ display: "none" }} />
      <input ref={arquivoImg} type="file" accept="image/png,image/jpeg,image/gif,image/bmp" onChange={inserirImagem} style={{ display: "none" }} />

      {/* Barra de título (como a do Word, nas cores do app) */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", background: COR_TITULO, color: "#fff", padding: "6px 10px", borderRadius: "10px 10px 0 0" }}
        onMouseDown={(e) => { if (e.target.closest("button")) e.preventDefault(); }}>
        <button style={sRapido} onClick={() => salvarDocx(false)} title={arquivoAtual ? `Salvar em "${arquivoAtual}" (Ctrl+S)` : "Salvar (Ctrl+S)"}>💾</button>
        <button style={sRapido} onClick={() => (atual.controlar ? desfazerControle() : cmd("undo"))} title="Desfazer (Ctrl+Z)">↶</button>
        <button style={sRapido} onClick={() => cmd("redo")} title="Refazer (Ctrl+Y)">↷</button>
        <span style={{ width: 1, height: 18, background: "rgba(255,255,255,0.3)" }} />
        <input value={atual.nome} onChange={(e) => renomear(e.target.value)} placeholder="Nome do documento" title="Nome do documento"
          style={{ flex: "1 1 200px", minWidth: 140, background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 4, color: "#fff", padding: "4px 8px", fontSize: 13, textAlign: "center", fontFamily: FONTE_UI }} />
        <span style={{ fontSize: 12, opacity: 0.9, whiteSpace: "nowrap", fontFamily: FONTE_UI }}>
          {arquivoAtual ? <>{arquivoAtual} · {atual.alterado ? "● não salvo" : "✓ salvo"}</> : suportaArquivoLocal ? "não salvo em arquivo" : "salva baixando cópia"}
        </span>
        <select value={atualId} onChange={(e) => trocar(e.target.value)} title="Documentos abertos neste navegador"
          style={{ maxWidth: 180, background: "rgba(255,255,255,0.12)", color: "#fff", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 4, fontSize: 12, padding: "3px 4px" }}>
          {docs.map((x) => <option key={x.id} value={x.id} style={{ color: "#222" }}>{x.nome || "(sem nome)"}</option>)}
        </select>
      </div>

      {/* Abas da faixa de opções */}
      <div style={{ display: "flex", gap: 2, background: COR_ABAS, padding: "0 6px", overflowX: "auto", fontFamily: FONTE_UI }}>
        {ABAS.map(([id, rot]) => (
          <button key={id} onMouseDown={(e) => e.preventDefault()} onClick={() => setAbaFaixa(id)}
            style={{ border: 0, padding: "7px 14px", fontSize: 13, cursor: "pointer", whiteSpace: "nowrap", fontFamily: FONTE_UI,
              background: id === "arquivo" ? GOLD : abaFaixa === id ? COR_FAIXA : "transparent", color: id === "arquivo" ? "#fff" : abaFaixa === id ? GOLD : "#3a3a3a",
              fontWeight: abaFaixa === id ? 600 : 400, borderRadius: "4px 4px 0 0", borderBottom: abaFaixa === id && id !== "arquivo" ? `2px solid ${GOLD}` : "2px solid transparent" }}>{rot}</button>
        ))}
      </div>

      {/* Faixa de opções */}
      <div style={{ display: "flex", alignItems: "stretch", flexWrap: "wrap", background: COR_FAIXA, borderBottom: "1px solid #d8d2c4", padding: "4px 4px 0", position: "sticky", top: 0, zIndex: 5, fontFamily: FONTE_UI, minHeight: 92 }}
        onMouseDown={(e) => { if (e.target.closest("button") || e.target.closest("[data-sem-foco]")) e.preventDefault(); }}>
        {abaFaixa === "arquivo" && <>
          <Grupo titulo="Documento">
            <Btn grande icone="＋" rotulo="Novo" onClick={() => criar()} />
            <Btn grande icone="📂" rotulo="Abrir" onClick={abrir} title="Abrir .docx, .txt ou .html do computador" />
            <Btn grande icone="💾" rotulo="Salvar" onClick={() => salvarDocx(false)} title={arquivoAtual ? `Gravar em "${arquivoAtual}" (Ctrl+S)` : "Salvar em Word (.docx) (Ctrl+S)"} />
            <Btn grande icone="🗂" rotulo="Salvar como" onClick={() => salvarDocx(true)} title="Salvar uma cópia em outro arquivo (Ctrl+Shift+S)" />
          </Grupo>
          <Grupo titulo="Imprimir">
            <Btn grande icone="📄" rotulo={gerandoPdf ? "Gerando…" : "Exportar PDF"} onClick={exportarPdf} title="Gera e baixa o PDF com as páginas como no Word (cabeçalho, rodapé, notas e números de página)" />
            <Btn grande icone="🖨" rotulo="Imprimir" onClick={() => imprimir()} title="Ctrl+P" />
          </Grupo>
          <Grupo titulo="Gerenciar">
            <Btn grande icone="🗑" rotulo="Excluir" onClick={excluir} title="Excluir este documento do navegador" />
          </Grupo>
          <div style={{ fontSize: 12, color: "#666", alignSelf: "center", padding: "0 12px", maxWidth: 420 }}>
            {arquivoAtual ? <>Ligado ao arquivo <b>{arquivoAtual}</b>: Salvar grava nele mesmo.</>
              : suportaArquivoLocal ? "Ainda não salvo em arquivo: Salvar pergunta onde guardar e, depois, grava sempre no mesmo arquivo."
              : "Este navegador não permite salvar no mesmo arquivo: Salvar baixa um .docx novo (use Chrome ou Edge no computador)."}
          </div>
        </>}

        {abaFaixa === "inicio" && <>
          <Grupo titulo="Área de Transferência">
            <Btn grande icone="📋" rotulo="Colar" onClick={colar} title="Colar como texto (Ctrl+V cola mantendo a formatação)" />
            <div style={sColuna}>
              <Btn icone="🖌" rotulo="Pincel" ativo={!!pincel} onClick={pegarFormato} title="Pincel de formatação: copia a formatação do trecho onde está o cursor para o próximo texto que você selecionar" />
              <Btn icone="✂" rotulo="Recortar" onClick={() => cmd("cut")} title="Ctrl+X" />
              <Btn icone="⧉" rotulo="Copiar" onClick={() => cmd("copy")} title="Ctrl+C" />
            </div>
          </Grupo>
          <Grupo titulo="Fonte">
            <div style={sColuna}>
              <div style={sLinha}>
                <select data-sem-foco style={{ ...sSelect, width: 132 }} defaultValue="" onChange={(e) => { cmd("fontName", e.target.value); e.target.value = ""; }} title="Fonte">
                  <option value="" disabled>{estilos.normal.fonte}</option>
                  {FONTES.map((f) => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
                </select>
                <select data-sem-foco style={{ ...sSelect, width: 56 }} defaultValue="" onChange={(e) => { aplicarTamanho(e.target.value); e.target.value = ""; }} title="Tamanho da fonte (pt)">
                  <option value="" disabled>{estilos.normal.tamanho}</option>
                  {TAMANHOS.map((pt) => <option key={pt} value={pt}>{String(pt).replace(".", ",")}</option>)}
                </select>
                <Btn icone="A⁺" onClick={() => tamanhoRelativo(1)} title="Aumentar fonte (Ctrl+])" />
                <Btn icone="A⁻" onClick={() => tamanhoRelativo(-1)} title="Diminuir fonte (Ctrl+[)" />
                <Menu icone="Aa" title="Maiúsculas e minúsculas" itens={[["Primeira letra da frase maiúscula", () => trocarCaixa("frase")], ["minúsculas", () => trocarCaixa("min")], ["MAIÚSCULAS", () => trocarCaixa("mai")], ["Colocar Cada Palavra Em Maiúscula", () => trocarCaixa("palavras")]]} />
                <Btn icone="⌫" onClick={() => cmd("removeFormat")} title="Limpar formatação" />
              </div>
              <div style={sLinha}>
                <Btn icone={<b>N</b>} ativo={estado.bold} onClick={() => cmd("bold")} title="Negrito (Ctrl+B)" />
                <Btn icone={<i>I</i>} ativo={estado.italic} onClick={() => cmd("italic")} title="Itálico (Ctrl+I)" />
                <Btn icone={<u>S</u>} ativo={estado.underline} onClick={() => cmd("underline")} title="Sublinhado (Ctrl+U)" />
                <Btn icone={<s>ab</s>} ativo={estado.strikeThrough} onClick={() => cmd("strikeThrough")} title="Tachado" />
                <Btn icone={<span>x<sub>2</sub></span>} ativo={estado.subscript} onClick={() => cmd("subscript")} title="Subscrito" />
                <Btn icone={<span>x<sup>2</sup></span>} ativo={estado.superscript} onClick={() => cmd("superscript")} title="Sobrescrito" />
                <label style={sCor} title="Cor da fonte">A<input type="color" onChange={(e) => cmd("foreColor", e.target.value)} style={sInputCor} /></label>
                <label style={sCor} title="Cor do realce do texto">🖍<input type="color" defaultValue="#ffff00" onChange={(e) => cmd("hiliteColor", e.target.value)} style={sInputCor} /></label>
              </div>
            </div>
          </Grupo>
          <Grupo titulo="Parágrafo">
            <div style={sColuna}>
              <div style={sLinha}>
                <Btn icone="•≡" ativo={estado.insertUnorderedList} onClick={() => cmd("insertUnorderedList")} title="Marcadores" />
                <Btn icone="1≡" ativo={estado.insertOrderedList} onClick={() => cmd("insertOrderedList")} title="Numeração" />
                <Menu title="Formato da numeração (Tab / Shift+Tab mudam o nível)" itens={[
                  ["1.  2.  3.", () => formatoNumeracao("decimal")],
                  ["1.  1.1.  1.1.1.  (cláusulas)", () => formatoNumeracao("juridico")],
                  ["1.  a)  i)", () => formatoNumeracao("alfa")],
                  ["I.  A.  1.", () => formatoNumeracao("romano")],
                  ["a)  b)  c)", () => formatoNumeracao("letras")],
                ]} />
                <Btn icone="⇤" onClick={() => cmd("outdent")} title="Diminuir recuo" />
                <Btn icone="⇥" onClick={() => cmd("indent")} title="Aumentar recuo" />
                <select data-sem-foco style={{ ...sSelect, width: 82 }} value="" onChange={(e) => estiloParagrafo("textIndent", e.target.value)} title="Recuo da primeira linha">
                  <option value="" disabled>Recuo ⇥</option>
                  <option value="0cm">Nenhum</option><option value="1.25cm">1,25 cm</option><option value="2.5cm">2,5 cm</option><option value="4cm">4 cm</option>
                </select>
              </div>
              <div style={sLinha}>
                <Btn icone="⯇≡" onClick={() => cmd("justifyLeft")} title="Alinhar à esquerda (Ctrl+L)" />
                <Btn icone="≡" onClick={() => cmd("justifyCenter")} title="Centralizar (Ctrl+E)" />
                <Btn icone="≡⯈" onClick={() => cmd("justifyRight")} title="Alinhar à direita (Ctrl+R)" />
                <Btn icone="☰" onClick={() => cmd("justifyFull")} title="Justificar (Ctrl+J)" />
                <select data-sem-foco style={{ ...sSelect, width: 62 }} value="" onChange={(e) => estiloParagrafo("lineHeight", e.target.value)} title="Espaçamento entre linhas">
                  <option value="" disabled>↕</option>
                  {["1", "1.15", "1.5", "2"].map((v) => <option key={v} value={v}>{v.replace(".", ",")}</option>)}
                </select>
                <Menu icone="▢" title="Bordas do parágrafo" itens={[["Borda inferior", () => bordaParagrafo("inferior")], ["Borda superior", () => bordaParagrafo("superior")], ["Caixa (todas as bordas)", () => bordaParagrafo("caixa")], ["Sem borda", () => bordaParagrafo("nenhuma")], ["Sem sombreamento", () => sombreamentoParagrafo("")]]} />
                <label style={sCor} title="Sombreamento do parágrafo">▤<input type="color" defaultValue="#eeeeee" onChange={(e) => sombreamentoParagrafo(e.target.value)} style={sInputCor} /></label>
                <select data-sem-foco style={{ ...sSelect, width: 82 }} value="" onChange={(e) => estiloParagrafo("marginBottom", e.target.value)} title="Espaço depois do parágrafo">
                  <option value="" disabled>¶ depois</option>
                  {["0pt", "6pt", "12pt", "18pt"].map((v) => <option key={v} value={v}>{v.replace("pt", " pt")}</option>)}
                </select>
              </div>
            </div>
          </Grupo>
          <Grupo titulo="Estilos">
            <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
              {[["p", "Normal", estilos.normal], ["h1", "Título 1", estilos.h1], ["h2", "Título 2", estilos.h2], ["h3", "Título 3", estilos.h3], ["blockquote", "Citação", estilos.citacao]].map(([tag, nome, e]) => {
                const ativo = (["h1", "h2", "h3", "blockquote"].includes(estado.bloco) ? estado.bloco : "p") === tag;
                return (
                  <button key={tag} onClick={() => cmd("formatBlock", tag)} title={`Estilo ${nome}`}
                    style={{ width: 56, height: 54, border: `1px solid ${ativo ? GOLD : "#ddd6c6"}`, background: ativo ? "#f3ead6" : "#fff", borderRadius: 3, cursor: "pointer", padding: 3, overflow: "hidden" }}>
                    <div style={{ fontFamily: `'${e.fonte}'`, fontWeight: e.negrito ? "bold" : "normal", fontStyle: e.italico ? "italic" : "normal", fontSize: Math.min(15, e.tamanho + 1), color: e.cor || "#222", lineHeight: 1.1 }}>AaBb</div>
                    <div style={{ fontSize: 10, color: "#555", marginTop: 4 }}>{nome}</div>
                  </button>
                );
              })}
              <Btn grande icone="⚙" rotulo="Modificar" ativo={painelEstilos} onClick={() => setPainelEstilos((x) => !x)} title="Editar os estilos Normal, Títulos e Citação" />
            </div>
          </Grupo>
          <Grupo titulo="Edição">
            <div style={sColuna}>
              <Btn icone="🔍" rotulo="Localizar" onClick={() => abrirBusca(false)} title="Ctrl+F" />
              <Btn icone="⇄" rotulo="Substituir" onClick={() => abrirBusca(true)} title="Ctrl+H" />
            </div>
          </Grupo>
        </>}

        {abaFaixa === "inserir" && <>
          <Grupo titulo="Páginas">
            <Btn grande icone="⤓" rotulo="Quebra de Página" onClick={() => cmd("insertHTML", "<hr><p><br></p>")} />
          </Grupo>
          <Grupo titulo="Tabelas">
            <span style={{ position: "relative" }}>
              <Btn grande icone="▦" rotulo="Tabela ▾" ativo={!!seletorTabela} onClick={() => setSeletorTabela((x) => (x ? null : { l: 0, c: 0 }))} />
              {seletorTabela && (
                <div style={{ position: "absolute", top: 64, left: 0, zIndex: 20, background: "#fff", border: "1px solid #ccc", borderRadius: 6, padding: 8, boxShadow: "0 4px 16px rgba(0,0,0,0.15)" }}
                  onMouseLeave={() => setSeletorTabela({ l: 0, c: 0 })}>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(10, 16px)", gap: 2 }}>
                    {Array.from({ length: 80 }, (_, i) => {
                      const l = Math.floor(i / 10) + 1, c = (i % 10) + 1;
                      const marc = l <= seletorTabela.l && c <= seletorTabela.c;
                      return <div key={i} onMouseEnter={() => setSeletorTabela({ l, c })} onMouseDown={(e) => e.preventDefault()} onClick={() => inserirTabela(l, c)}
                        style={{ width: 16, height: 16, border: "1px solid " + (marc ? GOLD : "#ccc"), background: marc ? "rgba(154,110,30,0.25)" : "#fff", cursor: "pointer" }} />;
                    })}
                  </div>
                  <div style={{ fontSize: 12, textAlign: "center", marginTop: 6, color: "#555" }}>
                    {seletorTabela.l ? `${seletorTabela.l} × ${seletorTabela.c}` : "Escolha o tamanho"}{" · "}
                    <a href="#" onMouseDown={(e) => e.preventDefault()} onClick={(e) => { e.preventDefault(); inserirTabela(parseInt(prompt("Número de linhas:", "3"), 10), parseInt(prompt("Número de colunas:", "3"), 10)); }}>outro tamanho…</a>
                  </div>
                </div>
              )}
            </span>
          </Grupo>
          <Grupo titulo="Ilustrações">
            <Btn grande icone="🖼" rotulo="Imagem" onClick={() => arquivoImg.current.click()} title="Inserir imagem do computador" />
          </Grupo>
          <Grupo titulo="Links">
            <Btn grande icone="🔗" rotulo="Link" onClick={inserirLink} />
          </Grupo>
          <Grupo titulo="Comentários">
            <Btn grande icone="💬" rotulo="Comentário" onClick={comentar} title="Comentar o trecho selecionado" />
          </Grupo>
          <Grupo titulo="Cabeçalho e Rodapé">
            <div style={sColuna}>
              <Btn icone="▔" rotulo="Cabeçalho" onClick={() => cab.current?.focus()} />
              <Btn icone="▁" rotulo="Rodapé" onClick={() => rod.current?.focus()} />
            </div>
            <select data-sem-foco style={{ ...sSelect, width: 132, alignSelf: "center" }} value={atual.numPag || ""} onChange={(e) => mudarDoc({ numPag: e.target.value })} title="Número de página">
              <option value="">Nº de página: não</option>
              <option value="direita">Rodapé, à direita</option>
              <option value="centro">Rodapé, centralizado</option>
              <option value="xdey">"Página X de Y"</option>
            </select>
            <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, alignSelf: "center", cursor: "pointer" }} title="Cabeçalho e rodapé próprios na primeira página (ex.: timbre só na folha de rosto). O número de página não aparece na 1ª página.">
              <input type="checkbox" checked={!!atual.primeiraDiferente} onChange={(e) => mudarDoc({ primeiraDiferente: e.target.checked })} />1ª página diferente
            </label>
          </Grupo>
          <Grupo titulo="Texto">
            <Menu grande icone="📅" rotulo="Data" title="Inserir a data de hoje" itens={[
              [textoData("extenso"), () => inserirSimbolo(textoData("extenso"))],
              [textoData("curta"), () => inserirSimbolo(textoData("curta"))],
              [textoData("semana"), () => inserirSimbolo(textoData("semana"))],
              [textoData("local"), () => inserirSimbolo(textoData("local"))],
            ]} />
          </Grupo>
          <Grupo titulo="Símbolos">
            <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 24px)", gap: 2 }}>
              {SIMBOLOS.map((c) => <button key={c} onClick={() => inserirSimbolo(c)} title={`Inserir ${c}`}
                style={{ width: 24, height: 24, border: "1px solid #e0dacb", background: "#fff", borderRadius: 3, cursor: "pointer", fontSize: 13, padding: 0 }}>{c}</button>)}
            </div>
          </Grupo>
        </>}

        {abaFaixa === "layout" && <>
          <Grupo titulo="Configurar Página">
            <Menu grande icone="▭" rotulo="Margens" itens={[
              ["Padrão ABNT (sup./esq. 3 cm; inf./dir. 2 cm)", () => mudarDoc({ pagina: { ...pagina, margens: { sup: 3, inf: 2, esq: 3, dir: 2 } } })],
              ["Normal (2,5 cm; esq./dir. 3 cm)", () => mudarDoc({ pagina: { ...pagina, margens: { sup: 2.5, inf: 2.5, esq: 3, dir: 3 } } })],
              ["Estreita (1,27 cm)", () => mudarDoc({ pagina: { ...pagina, margens: { sup: 1.27, inf: 1.27, esq: 1.27, dir: 1.27 } } })],
              ["Moderada (2,54 / 1,91 cm)", () => mudarDoc({ pagina: { ...pagina, margens: { sup: 2.54, inf: 2.54, esq: 1.91, dir: 1.91 } } })],
              ["Margens personalizadas…", () => setPainelPagina(true)],
            ]} />
            <Menu grande icone="⟳" rotulo="Orientação" itens={[["Retrato", () => mudarDoc({ pagina: { ...pagina, orient: "retrato" } })], ["Paisagem", () => mudarDoc({ pagina: { ...pagina, orient: "paisagem" } })]]} />
            <Menu grande icone="📄" rotulo="Tamanho" itens={[["A4 (21 × 29,7 cm)", () => mudarDoc({ pagina: { ...pagina, tam: "A4" } })], ["Ofício (21,6 × 33 cm)", () => mudarDoc({ pagina: { ...pagina, tam: "Oficio" } })], ["Carta (21,6 × 27,9 cm)", () => mudarDoc({ pagina: { ...pagina, tam: "Carta" } })]]} />
            <Menu grande icone="⤓" rotulo="Quebras" title="Quebra de página ou de seção" itens={[
              ["Quebra de página", () => cmd("insertHTML", "<hr><p><br></p>")],
              ["Quebra de seção: próximas páginas em paisagem", () => inserirQuebraSecao("paisagem")],
              ["Quebra de seção: próximas páginas em retrato", () => inserirQuebraSecao("retrato")],
            ]} />
          </Grupo>
          <Grupo titulo="Parágrafo">
            <div style={sColuna}>
              <div style={sLinha}><span style={sRotulo}>Recuo 1ª linha</span>
                <select data-sem-foco style={{ ...sSelect, width: 90 }} value="" onChange={(e) => estiloParagrafo("textIndent", e.target.value)}>
                  <option value="" disabled>escolher</option><option value="0cm">Nenhum</option><option value="1.25cm">1,25 cm</option><option value="2.5cm">2,5 cm</option><option value="4cm">4 cm</option>
                </select></div>
              <div style={sLinha}><span style={sRotulo}>Espaço depois</span>
                <select data-sem-foco style={{ ...sSelect, width: 90 }} value="" onChange={(e) => estiloParagrafo("marginBottom", e.target.value)}>
                  <option value="" disabled>escolher</option>{["0pt", "6pt", "12pt", "18pt"].map((v) => <option key={v} value={v}>{v.replace("pt", " pt")}</option>)}
                </select></div>
            </div>
          </Grupo>
          <Grupo titulo="Página atual">
            <div style={{ fontSize: 12, color: "#555", alignSelf: "center", lineHeight: 1.6 }}>
              {({ A4: "A4", Oficio: "Ofício", Carta: "Carta" })[pagina.tam]} · {pagina.orient === "paisagem" ? "paisagem" : "retrato"}<br />
              Margens: {m.sup} / {m.inf} / {m.esq} / {m.dir} cm
            </div>
          </Grupo>
        </>}

        {abaFaixa === "referencias" && <>
          <Grupo titulo="Sumário">
            <Btn grande icone="≣" rotulo="Sumário" onClick={inserirSumario} title="Inserir sumário automático (usa os Títulos 1, 2 e 3)" />
            <Btn grande icone="⟳" rotulo="Atualizar" onClick={() => { atualizarSumarios(); salvarHtml(); }} title="Atualizar o sumário" />
          </Grupo>
          <Grupo titulo="Notas de Rodapé">
            <Btn grande icone="¹" rotulo="Inserir Nota" onClick={inserirNota} title="Inserir nota de rodapé no ponto do cursor" />
          </Grupo>
        </>}

        {abaFaixa === "revisao" && <>
          <Grupo titulo="Revisão de Texto">
            <Btn grande icone="📊" rotulo="Contar Palavras" onClick={() => setContagem(true)} />
          </Grupo>
          <Grupo titulo="Comentários">
            <Btn grande icone="💬" rotulo="Novo Comentário" onClick={comentar} title="Comentar o trecho selecionado" />
          </Grupo>
          <Grupo titulo="Controle">
            <Btn grande icone="✎" rotulo={atual.controlar ? "Controlando: ligado" : "Controlar Alterações"} ativo={atual.controlar} onClick={alternarControle} title="Inclusões e exclusões ficam marcadas com o seu nome" />
            <div style={sColuna}>
              <label style={{ fontSize: 12, display: "flex", gap: 6, alignItems: "center" }}>
                <input type="checkbox" checked={mostrarMarcas} onChange={(e) => setMostrarMarcas(e.target.checked)} /> Mostrar marcações
              </label>
              <Btn icone="👤" rotulo={autor || "Meu nome"} onClick={() => { const n = prompt("Seu nome (aparece nos comentários e nas alterações):", autor)?.trim(); if (n) { gravar("editorTexto.autor", n); setAutor(n); } }} title="Nome usado nos comentários e alterações" />
            </div>
          </Grupo>
          <Grupo titulo="Alterações">
            <Btn grande icone="✓" rotulo="Aceitar Todas" onClick={() => revisarTodas(true)} />
            <Btn grande icone="✕" rotulo="Rejeitar Todas" onClick={() => revisarTodas(false)} />
          </Grupo>
        </>}

        {abaFaixa === "exibir" && <>
          <Grupo titulo="Zoom">
            <Btn grande icone="🔍" rotulo="100%" onClick={() => setZoom(100)} />
            <Btn grande icone="↔" rotulo="Largura da Página" onClick={larguraDaPagina} />
            <div style={sColuna}>
              <Btn icone="＋" rotulo="Aumentar" onClick={() => setZoom((z) => Math.min(200, z + 10))} />
              <Btn icone="－" rotulo="Diminuir" onClick={() => setZoom((z) => Math.max(50, z - 10))} />
            </div>
          </Grupo>
          <Grupo titulo="Mostrar">
            <div style={sColuna}>
              <label style={{ fontSize: 12, display: "flex", gap: 6, alignItems: "center" }}><input type="checkbox" checked={mostrarRegua} onChange={(e) => { setMostrarRegua(e.target.checked); gravar("editorTexto.regua", e.target.checked); }} /> Régua</label>
              <label style={{ fontSize: 12, display: "flex", gap: 6, alignItems: "center" }}><input type="checkbox" checked={mostrarQuebras} onChange={(e) => setMostrarQuebras(e.target.checked)} /> Marcas de fim de página</label>
              <label style={{ fontSize: 12, display: "flex", gap: 6, alignItems: "center" }}><input type="checkbox" checked={mostrarGrade} onChange={(e) => setMostrarGrade(e.target.checked)} /> Linhas de grade das tabelas</label>
              <label style={{ fontSize: 12, display: "flex", gap: 6, alignItems: "center" }}><input type="checkbox" checked={mostrarMarcas} onChange={(e) => setMostrarMarcas(e.target.checked)} /> Marcações de revisão</label>
            </div>
          </Grupo>
          <Grupo titulo="Fundo">
            <div style={sColuna}>
              <label style={{ fontSize: 12, display: "flex", gap: 6, alignItems: "center" }}>
                <input type="checkbox" checked={fundoEscuro} onChange={(e) => { setFundoEscuro(e.target.checked); gravar("editorTexto.fundoEscuro", e.target.checked); }} /> Fundo escuro em volta da página
              </label>
            </div>
          </Grupo>
        </>}
        {mostrarRegua && <Regua editor={editor} pagina={pagRef} area={areaRef} largCm={largCm} m={m} zoom={zoom / 100} versao={atual.html} blocos={blocosSelecionados} salvar={salvarHtml} ajustarTabs={ajustarTabs} />}
      </div>
      {msg && <p style={{ fontSize: 13, color: GOLD, margin: 0, padding: "6px 10px", background: "#fffaf0", borderBottom: "1px solid #eee", fontFamily: FONTE_UI }}>{msg}</p>}

      {busca && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center", padding: 8, background: "#fbf8f1", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0 }}>
          <input ref={campoBusca} placeholder="Localizar" value={busca.termo} onChange={(e) => { destacar(null); setBusca({ ...busca, termo: e.target.value, info: "" }); }}
            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); localizar(e.shiftKey); } if (e.key === "Escape") { e.preventDefault(); fecharBusca(); } }}
            style={{ ...sCampo, width: 200 }} />
          <input placeholder="Substituir por" value={busca.novo} onChange={(e) => setBusca({ ...busca, novo: e.target.value })}
            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); substituir(); } if (e.key === "Escape") { e.preventDefault(); fecharBusca(); } }}
            style={{ ...sCampo, width: 200 }} />
          <label style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 4 }}>
            <input type="checkbox" checked={busca.caso} onChange={(e) => setBusca({ ...busca, caso: e.target.checked, info: "" })} /> Diferenciar maiúsculas
          </label>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={() => localizar(true)} disabled={!busca.termo} title="Anterior (Shift+Enter)">↑</button>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={() => localizar(false)} disabled={!busca.termo} title="Próximo (Enter)">↓ Próximo</button>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={substituir} disabled={!busca.termo}>Substituir</button>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={substituirTudo} disabled={!busca.termo}>Substituir tudo</button>
          <span style={{ fontSize: 12, color: "#666" }}>{busca.info}</span>
          <button style={{ ...sBotao(), marginLeft: "auto" }} onClick={fecharBusca} title="Fechar (Esc)">×</button>
        </div>
      )}

      {painelEstilos && <PainelEstilos estilos={estilos} mudar={(e) => mudarDoc({ estilos: e })} />}

      {painelPagina && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", padding: 10, background: "#fbf8f1", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0, fontSize: 13 }}>
          <label>Papel{" "}
            <select style={sSelect} value={pagina.tam} onChange={(e) => mudarDoc({ pagina: { ...pagina, tam: e.target.value } })}>
              <option value="A4">A4 (21 × 29,7 cm)</option>
              <option value="Oficio">Ofício (21,6 × 33 cm)</option>
              <option value="Carta">Carta (21,6 × 27,9 cm)</option>
            </select>
          </label>
          <label>Orientação{" "}
            <select style={sSelect} value={pagina.orient} onChange={(e) => mudarDoc({ pagina: { ...pagina, orient: e.target.value } })}>
              <option value="retrato">Retrato</option>
              <option value="paisagem">Paisagem</option>
            </select>
          </label>
          <span>Margens (cm):</span>
          {[["sup", "Superior"], ["inf", "Inferior"], ["esq", "Esquerda"], ["dir", "Direita"]].map(([k, rot]) => (
            <label key={k}>{rot}{" "}
              <input type="number" step="0.25" min="0" max="10" value={m[k]} style={{ ...sCampo, width: 64 }}
                onChange={(e) => mudarDoc({ pagina: { ...pagina, margens: { ...m, [k]: Math.max(0, Number(e.target.value) || 0) } } })} />
            </label>
          ))}
          <button style={sBotao()} onClick={() => mudarDoc({ pagina: PAGINA_PADRAO })} title="Superior e esquerda 3 cm; inferior e direita 2 cm">Padrão ABNT</button>
          <label>Número de página{" "}
            <select style={sSelect} value={atual.numPag || ""} onChange={(e) => mudarDoc({ numPag: e.target.value })}>
              <option value="">Sem número</option>
              <option value="direita">No rodapé, à direita</option>
              <option value="centro">No rodapé, centralizado</option>
              <option value="xdey">"Página X de Y", centralizado</option>
            </select>
          </label>
          <span style={{ color: "#777" }}>O cabeçalho e o rodapé são escritos direto na folha, acima e abaixo do texto, e se repetem em todas as páginas.</span>
        </div>
      )}

      <div ref={areaRef} style={{ background: fundoEscuro ? COR_FUNDO_ESCURO : "#e9e6df", padding: "24px 12px", overflowX: "auto" }}>
      <div style={{ zoom: zoom / 100, display: "flex", justifyContent: "safe center", alignItems: "flex-start", gap: 14 }}>
        <div ref={pagRef} style={{ position: "relative", background: "#fff", width: `${largCm}cm`, maxWidth: "100%", flex: "none", marginRight: temSecaoLarga ? `${deltaSec}cm` : 0, minHeight: `${altCm}cm`, boxShadow: "0 2px 12px rgba(0,0,0,0.15)", boxSizing: "border-box", padding: `0 ${m.dir}cm 0 ${m.esq}cm`, display: "flex", flexDirection: "column" }}>
          <div style={{ minHeight: `${m.sup}cm`, display: "flex", alignItems: "flex-end", paddingBottom: "0.3cm", boxSizing: "border-box" }}>
            <div ref={cab} className="ed-area" contentEditable suppressContentEditableWarning data-vazio="Cabeçalho (clique para escrever)" lang="pt-BR" spellCheck
              onInput={() => mudarDoc({ cabecalho: cab.current.innerHTML })} style={{ width: "100%", fontFamily: `'${estilos.normal.fonte}', sans-serif` }} />
          </div>
          <div style={{ position: "relative", flex: 1 }}>
            {faixas.map((f, i) => (
              <div key={i} title="Seção em paisagem" style={{ position: "absolute", left: `calc(100% + ${m.dir}cm)`, width: `${deltaSec}cm`, top: f.topo, height: Math.max(0, f.fim - f.topo), background: "#fff", boxShadow: "6px 2px 10px rgba(0,0,0,0.15)", pointerEvents: "none" }} />
            ))}
            <div ref={editor} className={"ed-texto" + (mostrarMarcas ? "" : " sem-marcas") + (pincel ? " com-pincel" : "")} contentEditable suppressContentEditableWarning spellCheck lang="pt-BR"
              onInput={salvarHtml} onClick={cliqueEditor} onMouseUp={soltarPincel} onMouseDown={mouseDownEditor} onMouseMove={mouseMoveEditor} onContextMenu={menuContexto}
              style={{ minHeight: alturaUtil, position: "relative" }} />
            {ordemNotas.length > 0 && (
              <div className="ed-notas" style={{ marginTop: 12, paddingTop: 4, borderTop: "1px solid #888", width: "100%", fontFamily: `'${estilos.normal.fonte}', serif` }}>
                {ordemNotas.map((id, i) => (
                  <div key={id} style={{ display: "flex", gap: 6, alignItems: "baseline", marginBottom: 2 }}>
                    <sup style={{ color: GOLD, fontWeight: "bold", cursor: "pointer" }} title="Ir para a chamada da nota"
                      onClick={() => editor.current.querySelector(`sup.nota[data-nota="${id}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" })}>{i + 1}</sup>
                    <NotaTexto id={id} html={atual.notas?.[id] || ""} docId={atual.id} mudar={mudarNota} />
                  </div>
                ))}
              </div>
            )}
            {imgSel?.isConnected && <MolduraImagem zoom={zoom / 100} img={imgSel} base={editor.current} redimensionar={redimensionarImagem} posicionar={posicionarImagem} original={tamanhoOriginal} excluir={excluirImagem} />}
            {mostrarQuebras && quebras.map((y, i) => (
              <div key={i} style={{ position: "absolute", left: `-${m.esq}cm`, right: `-${m.dir}cm`, top: y, borderTop: "1px dashed #c9bfa8", pointerEvents: "none" }}>
                <span style={{ position: "absolute", right: 4, top: -9, fontSize: 10, color: "#a8987a", background: "#fff", padding: "0 4px", fontFamily: "sans-serif" }}>Página {i + 2}</span>
              </div>
            ))}
          </div>
          <div style={{ minHeight: `${m.inf}cm`, paddingTop: "0.3cm", boxSizing: "border-box" }}>
            <div ref={rod} className="ed-area" contentEditable suppressContentEditableWarning data-vazio="Rodapé (clique para escrever)" lang="pt-BR" spellCheck
              onInput={() => mudarDoc({ rodape: rod.current.innerHTML })} style={{ width: "100%", fontFamily: `'${estilos.normal.fonte}', sans-serif` }} />
            {textoNumero && (
              <div style={{ fontSize: "10pt", fontFamily: `'${estilos.normal.fonte}', sans-serif`, textAlign: atual.numPag === "direita" ? "right" : "center", color: "#555", padding: "0 4px" }} title="Número de página (aparece em todas as páginas no Word e na impressão)">
                {textoNumero}
              </div>
            )}
          </div>
        </div>
        {atual.primeiraDiferente && (
          <div style={{ width: 250, flex: "none", background: "#fbf8f1", border: "1px solid rgba(0,0,0,0.15)", borderRadius: 6, padding: 10, fontFamily: FONTE_UI, fontSize: 12, boxShadow: "0 2px 8px rgba(0,0,0,0.12)" }}>
            <div style={{ fontWeight: 600, color: COR_TITULO, marginBottom: 6 }}>Primeira página diferente</div>
            <div style={{ color: "#666", marginBottom: 8, lineHeight: 1.35 }}>Na página o cabeçalho e o rodapé são os das <b>demais páginas</b>. Aqui ficam os da <b>1ª página</b> (em branco = sem cabeçalho/rodapé nela).</div>
            {[["cabecalho1", "Cabeçalho da 1ª página"], ["rodape1", "Rodapé da 1ª página"]].map(([campo, rotulo]) => (
              <div key={campo} style={{ marginBottom: 8 }}>
                <div style={{ color: "#555", marginBottom: 3 }}>{rotulo}</div>
                <AreaPrimeira campo={campo} html={atual[campo] || ""} docId={atual.id} mudar={(v) => mudarDoc({ [campo]: v })} fonte={estilos.normal.fonte} />
              </div>
            ))}
          </div>
        )}
        {(Object.keys(comentarios).length > 0 || autoresRev.length > 0 || atual.controlar) && (
          <PainelRevisao zoom={zoom / 100} editor={editor.current} pagina={pagRef.current} comentarios={comentarios} versao={atual.html} quebras={quebras}
            ativo={comAtivo} setAtivo={setComAtivo} novo={comNovo} setNovo={setComNovo} autor={autor} mostrarMarcas={mostrarMarcas}
            mudar={mudarComentario} excluir={excluirComentario} revisar={revisar} />
        )}
      </div>
      </div>
      {/* Barra de status */}
      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap", fontSize: 12, color: "#e8e4da", background: COR_STATUS, padding: "4px 12px", borderRadius: "0 0 10px 10px", fontFamily: FONTE_UI }}>
        <span>Página {Math.min(pagAtual, totalPag)} de {totalPag}</span>
        <button onClick={() => setContagem(true)} style={{ background: "none", border: 0, color: "inherit", cursor: "pointer", fontSize: 12, padding: 0 }} title="Contar palavras">{palavras} palavra{palavras === 1 ? "" : "s"}</button>
        <span>Português (Brasil)</span>
        {atual.controlar && <span style={{ color: "#ffd38a" }}>✎ Controlando alterações</span>}
        <span style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6 }}>
          <button onClick={() => setZoom((z) => Math.max(50, z - 10))} style={sZoomBt} title="Diminuir zoom">－</button>
          <input type="range" min="50" max="200" step="10" value={zoom} onChange={(e) => setZoom(Number(e.target.value))} style={{ width: 110 }} title="Zoom" />
          <button onClick={() => setZoom((z) => Math.min(200, z + 10))} style={sZoomBt} title="Aumentar zoom">＋</button>
          <button onClick={() => setZoom(100)} style={{ ...sZoomBt, width: 44 }} title="Voltar a 100%">{zoom}%</button>
        </span>
      </div>
      {contagem && <ContagemPalavras editor={editor.current} paginas={totalPag} fechar={() => setContagem(false)} />}
      {menu && <MenuTabela menu={menu} sel={selCelulas.current?.ret} op={opTabela} fechar={() => setMenu(null)} />}
    </div>
  );
}

function MenuTabela({ menu, sel, op, fechar }) {
  useEffect(() => {
    const fora = (e) => { if (!e.target.closest?.("[data-menu-tabela]")) fechar(); };
    const esc = (e) => { if (e.key === "Escape") fechar(); };
    window.addEventListener("mousedown", fora);
    window.addEventListener("keydown", esc);
    return () => { window.removeEventListener("mousedown", fora); window.removeEventListener("keydown", esc); };
  }, [fechar]);
  const td = menu.td;
  const varias = sel && sel.celulas.length > 1;
  const mesclada = td.rowSpan > 1 || td.colSpan > 1;
  const item = { display: "block", width: "100%", textAlign: "left", padding: "6px 12px", border: 0, background: "none", fontSize: 13, cursor: "pointer", fontFamily: "sans-serif" };
  const sep = <div style={{ borderTop: "1px solid #eee", margin: "4px 0" }} />;
  const I = ({ children, f, desab }) => (
    <button style={{ ...item, color: desab ? "#bbb" : "#222" }} disabled={desab} onMouseDown={(e) => e.preventDefault()}
      onMouseEnter={(e) => { e.currentTarget.style.background = "#f4eee2"; }} onMouseLeave={(e) => { e.currentTarget.style.background = "none"; }}
      onClick={() => op(f)}>{children}</button>
  );
  return (
    <div data-menu-tabela style={{ position: "fixed", left: menu.x, top: menu.y, zIndex: 60, width: 240, background: "#fff", border: "1px solid rgba(0,0,0,0.2)", borderRadius: 8, boxShadow: "0 6px 24px rgba(0,0,0,0.2)", padding: "4px 0" }}>
      <I f={(t, c) => T.inserirLinha(t, c, false)}>↥ Inserir linha acima</I>
      <I f={(t, c) => T.inserirLinha(t, c, true)}>↧ Inserir linha abaixo</I>
      <I f={(t, c) => T.inserirColuna(t, c, false)}>↤ Inserir coluna à esquerda</I>
      <I f={(t, c) => T.inserirColuna(t, c, true)}>↦ Inserir coluna à direita</I>
      {sep}
      <I f={(t, c) => T.excluirLinhas(t, c)}>✕ Excluir linha</I>
      <I f={(t, c) => T.excluirColunas(t, c)}>✕ Excluir coluna</I>
      <I f={(t) => t.remove()}>✕ Excluir tabela</I>
      {sep}
      <I f={(t, c, s) => T.mesclar(t, s)} desab={!varias}>⊞ Mesclar células{varias ? "" : " (arraste para selecionar)"}</I>
      <I f={(t, c) => T.dividir(t, c)} desab={!mesclada}>⊟ Dividir célula</I>
      {sep}
      <I f={(t, c, s) => T.bordas(t, s?.celulas, "todas")}>▦ Todas as bordas</I>
      <I f={(t, c, s) => T.bordas(t, s?.celulas, "externas")}>▢ Só bordas externas</I>
      <I f={(t, c, s) => T.bordas(t, s?.celulas, "nenhuma")}>⬚ Sem bordas</I>
      <div style={{ ...item, display: "flex", alignItems: "center", gap: 8, cursor: "default" }}>
        🎨 Sombreamento
        <input type="color" defaultValue="#f2f2f2" onChange={(e) => { const cor = e.target.value; op((t, c, s) => (s?.celulas.length ? s.celulas : [c]).forEach((x) => { x.style.backgroundColor = cor; })); }} />
        <button style={{ ...item, width: "auto", padding: "2px 6px", border: "1px solid #ddd", borderRadius: 4 }} onMouseDown={(e) => e.preventDefault()}
          onClick={() => op((t, c, s) => (s?.celulas.length ? s.celulas : [c]).forEach((x) => { x.style.backgroundColor = ""; }))}>sem cor</button>
      </div>
    </div>
  );
}

function MolduraImagem({ zoom = 1, img, base, redimensionar, posicionar, original, excluir }) {
  const b = base.getBoundingClientRect(), r0 = img.getBoundingClientRect();
  const r = { width: r0.width / zoom, height: r0.height / zoom };
  const x = (r0.left - b.left) / zoom, y = (r0.top - b.top) / zoom;
  const flutua = img.style.float;
  const centro = img.closest("p,div,h1,h2,h3,li,td")?.style.textAlign === "center";
  const alca = (canto, estilo) => (
    <div key={canto} onMouseDown={(e) => redimensionar(e, canto)}
      style={{ position: "absolute", width: 10, height: 10, background: "#fff", border: "1.5px solid #3b6fd6", borderRadius: 2, cursor: canto === "no" || canto === "sd" ? "nwse-resize" : "nesw-resize", ...estilo }} />
  );
  const bt = (ativo) => ({ border: "1px solid #ddd", background: ativo ? "rgba(154,110,30,0.2)" : "#fff", borderRadius: 4, padding: "2px 7px", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" });
  return (
    <div data-moldura style={{ position: "absolute", left: x, top: y, width: r.width, height: r.height, outline: "1.5px solid #3b6fd6", pointerEvents: "none", zIndex: 4 }}>
      <div style={{ pointerEvents: "auto" }}>
        {alca("no", { left: -6, top: -6 })}{alca("ne", { right: -6, top: -6 })}
        {alca("so", { left: -6, bottom: -6 })}{alca("sd", { right: -6, bottom: -6 })}
      </div>
      <div onMouseDown={(e) => e.preventDefault()} style={{ position: "absolute", left: 0, top: y > 40 ? -34 : r.height + 6, display: "flex", gap: 4, pointerEvents: "auto", background: "#fff", padding: 3, borderRadius: 6, boxShadow: "0 2px 10px rgba(0,0,0,0.18)", fontFamily: "sans-serif" }}>
        <button style={bt(!flutua && !centro)} onClick={() => posicionar("linha")} title="Imagem na linha do texto">Em linha</button>
        <button style={bt(centro && !flutua)} onClick={() => posicionar("centro")} title="Centralizar">Centro</button>
        <button style={bt(flutua === "left")} onClick={() => posicionar("esq")} title="Imagem à esquerda, texto ao redor">◧ Texto ao redor</button>
        <button style={bt(flutua === "right")} onClick={() => posicionar("dir")} title="Imagem à direita, texto ao redor">Texto ao redor ◨</button>
        <button style={bt()} onClick={original} title="Voltar ao tamanho original (limitado à largura da página)">Tamanho original</button>
        <button style={bt()} onClick={excluir} title="Excluir imagem (Delete)">🗑</button>
        <span style={{ fontSize: 11, color: "#888", alignSelf: "center", padding: "0 4px" }}>{Math.round(r.width)}×{Math.round(r.height)}</span>
      </div>
    </div>
  );
}

// Texto de uma nota de rodapé (editável; o HTML só é recarregado ao trocar de nota/documento).
// ---------- Régua (como a do Word) ----------
// Recuo da 1ª linha (▽ em cima), recuo deslocado (△ embaixo), recuo esquerdo (▭, move os
// dois juntos) e recuo direito (△ à direita). Clique na régua cria uma tabulação; arraste
// a tabulação para fora da régua para apagá-la.
function Regua({ editor, pagina, area, largCm, m, zoom, versao, blocos, salvar, ajustarTabs }) {
  const [info, setInfo] = useState(null);
  const [esquerda, setEsquerda] = useState(0); // onde a folha começa, em px dentro da barra
  const ref = useRef(null);
  const barra = useRef(null);
  // Fica presa embaixo da faixa (como no Word) e acompanha a folha na horizontal.
  useLayoutEffect(() => {
    const alinhar = () => {
      if (!pagina.current || !barra.current) return;
      setEsquerda(pagina.current.getBoundingClientRect().left - barra.current.getBoundingClientRect().left);
    };
    alinhar();
    const ro = new ResizeObserver(alinhar);
    if (pagina.current) ro.observe(pagina.current);
    if (barra.current) ro.observe(barra.current);
    const el = area.current;
    el?.addEventListener("scroll", alinhar);
    window.addEventListener("resize", alinhar);
    return () => { ro.disconnect(); el?.removeEventListener("scroll", alinhar); window.removeEventListener("resize", alinhar); };
  });
  const medir = () => {
    const ed = editor.current;
    const s = document.getSelection();
    if (!ed || !s.rangeCount || !ed.contains(s.anchorNode)) return;
    const n = s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode;
    const b = n.closest("p,h1,h2,h3,h4,h5,h6,li,blockquote,div:not(.ed-texto)");
    if (!b || !ed.contains(b)) return setInfo(null);
    const z = zoom;
    const er = ed.getBoundingClientRect(), br = b.getBoundingClientRect();
    const cel = b.closest("td,th");
    const esq = m.esq + (br.left - er.left) / z / CM;
    const dir = largCm - m.dir - (er.right - br.right) / z / CM;
    const ti = parseFloat(getComputedStyle(b).textIndent) / CM || 0;
    const origem = cel ? m.esq + ((cel.getBoundingClientRect().left - er.left) / z + (parseFloat(getComputedStyle(cel).paddingLeft) || 0)) / CM : m.esq;
    const tabs = (b.dataset.tabs || "").split(",").map(Number).filter((v) => v > 0);
    setInfo({ esq, dir, primeira: esq + ti, origem, tabs });
  };
  useEffect(() => {
    document.addEventListener("selectionchange", medir);
    return () => document.removeEventListener("selectionchange", medir);
  });
  useEffect(() => { medir(); }, [versao, zoom, largCm, m.esq, m.dir]);

  const cmDoMouse = (ev, livre) => {
    const r = ref.current.getBoundingClientRect();
    const v = (ev.clientX - r.left) / zoom / CM;
    return Math.max(0, Math.min(largCm, livre ? Math.round(v * 100) / 100 : Math.round(v * 4) / 4)); // passo de 0,25 cm (Alt: livre)
  };
  const fmt = (v) => `${Math.round(v * 100) / 100}cm`;

  function arrastar(e, tipo, tabIdx) {
    e.preventDefault();
    e.stopPropagation();
    if (!info) return;
    const alvo = blocos();
    if (!alvo.length) return;
    const ini = { ...info, tabs: [...info.tabs] };
    // Recuos atuais de cada bloco, em cm (lidos do estilo calculado).
    const base = alvo.map((b) => ({ b }));
    base.forEach((x) => {
      const cs = getComputedStyle(x.b);
      x.ml = parseFloat(cs.marginLeft) / CM || 0;
      x.mr = parseFloat(cs.marginRight) / CM || 0;
      x.ti = parseFloat(cs.textIndent) / CM || 0;
    });
    let mudou = false;
    const mover = (ev) => {
      const v = cmDoMouse(ev, ev.altKey);
      mudou = true;
      if (tipo === "primeira") base.forEach((x) => { x.b.style.textIndent = fmt(v - ini.esq + 0); });
      if (tipo === "deslocado") base.forEach((x) => { // mexe o esquerdo e mantém a 1ª linha no lugar
        const d = v - ini.esq;
        x.b.style.marginLeft = fmt(Math.max(-m.esq, x.ml + d));
        x.b.style.textIndent = fmt(x.ti - d);
      });
      if (tipo === "esquerdo") base.forEach((x) => { x.b.style.marginLeft = fmt(Math.max(-m.esq, x.ml + v - ini.esq)); });
      if (tipo === "direito") base.forEach((x) => { x.b.style.marginRight = fmt(Math.max(-m.dir, x.mr + ini.dir - v)); });
      if (tipo === "tab") {
        const fora = Math.abs(ev.clientY - ref.current.getBoundingClientRect().top - 9) > 28;
        const lista = ini.tabs.filter((_, i) => i !== tabIdx);
        if (!fora) lista.push(Math.round((v - ini.origem) * 100) / 100);
        alvo.forEach((b) => { if (lista.length) b.dataset.tabs = lista.filter((t) => t > 0).sort((a, c) => a - c).join(","); else b.removeAttribute("data-tabs"); });
      }
      ajustarTabs();
      medir();
    };
    const soltar = () => {
      window.removeEventListener("mousemove", mover);
      window.removeEventListener("mouseup", soltar);
      if (mudou) salvar();
    };
    window.addEventListener("mousemove", mover);
    window.addEventListener("mouseup", soltar);
  }

  function novaTab(e) {
    e.preventDefault();
    if (!info) return;
    const v = cmDoMouse(e, e.altKey);
    if (v <= info.origem || v >= largCm - m.dir) return;
    const t = Math.round((v - info.origem) * 100) / 100;
    blocos().forEach((b) => {
      const lista = (b.dataset.tabs || "").split(",").map(Number).filter((x) => x > 0 && Math.abs(x - t) > 0.05);
      b.dataset.tabs = [...lista, t].sort((a, c) => a - c).join(",");
    });
    ajustarTabs();
    salvar();
    medir();
  }

  const marcas = [];
  for (let q = 0; q <= largCm * 4; q++) {
    const x = q / 4;
    const rel = x - m.esq;
    const inteiro = Math.abs(rel - Math.round(rel)) < 0.01;
    const meio = Math.abs(Math.abs(rel % 1) - 0.5) < 0.01;
    if (inteiro && Math.round(rel) !== 0) marcas.push(<span key={q} style={{ position: "absolute", left: `${x}cm`, top: 3, transform: "translateX(-50%)", fontSize: 9, color: "#555", fontFamily: "sans-serif", lineHeight: 1 }}>{Math.abs(Math.round(rel))}</span>);
    else if (!inteiro) marcas.push(<span key={q} style={{ position: "absolute", left: `${x}cm`, top: meio ? 6 : 8, height: meio ? 6 : 2, borderLeft: "1px solid #777" }} />);
  }
  const tri = (cima) => <svg width="11" height="8" style={{ display: "block" }}><path d={cima ? "M0.5 0.5 H10.5 L5.5 7.5 Z" : "M0.5 7.5 H10.5 L5.5 0.5 Z"} fill="#e9e6df" stroke="#555" /></svg>;
  const marcador = (x, top, conteudo, tipo, title) => (
    <div title={title} onMouseDown={(e) => arrastar(e, tipo)} style={{ position: "absolute", left: `${x}cm`, top, transform: "translateX(-50%)", cursor: "ew-resize", lineHeight: 0, zIndex: 2 }}>{conteudo}</div>
  );
  return (
    <div ref={barra} data-sem-foco style={{ flexBasis: "100%", height: Math.round(18 * zoom + 8), position: "relative", margin: "4px -4px 0", background: "#dedad0", borderTop: "1px solid #d0c9b8", overflow: "hidden" }}>
    <div style={{ position: "absolute", left: esquerda / zoom, top: 4 / zoom, zoom }}>
    <div ref={ref} data-regua onMouseDown={novaTab} title="Clique para criar uma tabulação"
      style={{ position: "relative", width: `${largCm}cm`, height: 18, background: "#fff", boxShadow: "0 1px 2px rgba(0,0,0,0.25)", cursor: "default", userSelect: "none" }}>
      <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${m.esq}cm`, background: "#cfcabd" }} />
      <div style={{ position: "absolute", right: 0, top: 0, bottom: 0, width: `${m.dir}cm`, background: "#cfcabd" }} />
      {marcas}
      {info && <>
        {marcador(info.primeira, -1, tri(true), "primeira", "Recuo da primeira linha")}
        {marcador(info.esq, 10, tri(false), "deslocado", "Recuo deslocado")}
        {marcador(info.esq, 18, <div style={{ width: 9, height: 5, background: "#e9e6df", border: "1px solid #555" }} />, "esquerdo", "Recuo à esquerda")}
        {marcador(info.dir, 10, tri(false), "direito", "Recuo à direita")}
        {info.tabs.map((t, i) => (
          <div key={i} title="Tabulação (arraste para fora da régua para apagar)" onMouseDown={(e) => arrastar(e, "tab", i)}
            style={{ position: "absolute", left: `${info.origem + t}cm`, top: 9, width: 7, height: 7, borderLeft: "2px solid #222", borderBottom: "2px solid #222", cursor: "ew-resize", zIndex: 2 }} />
        ))}
      </>}
    </div>
    </div>
    </div>
  );
}

function AreaPrimeira({ campo, html, docId, mudar, fonte }) {
  const ref = useRef(null);
  useEffect(() => { if (ref.current && ref.current.innerHTML !== html) ref.current.innerHTML = html; }, [campo, docId]);
  return <div ref={ref} className="ed-area" data-area={campo} contentEditable suppressContentEditableWarning spellCheck lang="pt-BR" data-vazio="(em branco)"
    onInput={() => mudar(ref.current.innerHTML)} style={{ background: "#fff", border: "1px solid #d0c8b6", minHeight: 40, padding: 6, fontSize: "11pt", fontFamily: `'${fonte}', sans-serif`, color: "#111" }} />;
}

function NotaTexto({ id, html, docId, mudar }) {
  const ref = useRef(null);
  useEffect(() => { if (ref.current && ref.current.innerHTML !== html) ref.current.innerHTML = html || "<p><br></p>"; }, [id, docId]);
  return <div ref={ref} data-nota-texto={id} contentEditable suppressContentEditableWarning spellCheck lang="pt-BR"
    onInput={() => mudar(id, ref.current.innerHTML)} style={{ flex: 1, outline: "none", minHeight: "1.2em" }} />;
}

const FONTES_ESTILO = ["Times New Roman", "Arial", "Arial Narrow", "Calibri", "Georgia", "Verdana", "Courier New", "Arial Nova Cond"];
function PainelEstilos({ estilos, mudar }) {
  const campo = { height: 26, border: "1px solid rgba(0,0,0,0.15)", borderRadius: 5, fontSize: 12, padding: "0 4px", background: "#fff" };
  const set = (k, prop, v) => mudar({ ...estilos, [k]: { ...estilos[k], [prop]: v } });
  const num = (k, prop, passo = 1, w = 52) => <input type="number" step={passo} min="0" value={estilos[k][prop] ?? 0} style={{ ...campo, width: w }} onChange={(e) => set(k, prop, Number(e.target.value) || 0)} />;
  const sel = (k, prop, opcoes) => (
    <select value={estilos[k][prop]} style={campo} onChange={(e) => set(k, prop, e.target.value)}>
      {opcoes.map(([v, t]) => <option key={v} value={v}>{t}</option>)}
    </select>
  );
  const alin = [["left", "Esquerda"], ["center", "Centro"], ["right", "Direita"], ["justify", "Justificado"]];
  const th = { textAlign: "left", fontWeight: 500, color: "#666", padding: "2px 6px", fontSize: 11 };
  const td = { padding: "3px 6px", whiteSpace: "nowrap" };
  return (
    <div style={{ padding: 10, background: "#fbf8f1", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0, fontSize: 12, overflowX: "auto" }}>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr>
          <th style={th}>Estilo</th><th style={th}>Fonte</th><th style={th}>Tam. (pt)</th><th style={th}>Alinhamento</th>
          <th style={th}>Antes (pt)</th><th style={th}>Depois (pt)</th><th style={th}>Entrelinhas</th><th style={th}>Recuo (cm)</th><th style={th}>Negrito / Itálico</th>
        </tr></thead>
        <tbody>
          {Object.keys(ESTILOS_PADRAO).map((k) => (
            <tr key={k} style={{ borderTop: "1px solid #eee" }}>
              <td style={{ ...td, fontWeight: 600 }}>{estilos[k].nome}</td>
              <td style={td}>{sel(k, "fonte", FONTES_ESTILO.map((f) => [f, f]))}</td>
              <td style={td}>{num(k, "tamanho", 0.5)}</td>
              <td style={td}>{sel(k, "alinhamento", alin)}</td>
              <td style={td}>{num(k, "antes")}</td>
              <td style={td}>{num(k, "depois")}</td>
              <td style={td}>{"entrelinhas" in estilos[k] ? sel(k, "entrelinhas", [[1, "1,0"], [1.15, "1,15"], [1.5, "1,5"], [2, "2,0"]].map(([v, t]) => [v, t])) : "—"}</td>
              <td style={td}>{k === "normal" ? <>1ª linha {num(k, "recuo", 0.25)}</> : k === "citacao" ? <>esquerda {num(k, "recuoEsq", 0.25)}</> : "—"}</td>
              <td style={td}>
                {"negrito" in estilos[k] && <label><input type="checkbox" checked={!!estilos[k].negrito} onChange={(e) => set(k, "negrito", e.target.checked)} /> N </label>}
                <label><input type="checkbox" checked={!!estilos[k].italico} onChange={(e) => set(k, "italico", e.target.checked)} /> I</label>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{ marginTop: 8, display: "flex", gap: 8, alignItems: "center" }}>
        <button style={{ ...campo, cursor: "pointer", padding: "0 10px" }} onClick={() => mudar(undefined)}>Restaurar padrão</button>
        <span style={{ color: "#777" }}>Os estilos valem para o documento todo (na tela, na impressão e no Word). Formatação aplicada direto num trecho continua valendo por cima do estilo.</span>
      </div>
    </div>
  );
}

// Coluna à direita da página com os balões de comentários e de alterações, cada um na
// altura do trecho a que se refere (como no Word).
function PainelRevisao({ zoom = 1, editor, pagina, comentarios, versao, quebras, ativo, setAtivo, novo, setNovo, autor, mostrarMarcas, mudar, excluir, revisar }) {
  const [tops, setTops] = useState({});
  const caixa = useRef(null);
  const itens = [];
  if (editor && pagina) {
    const base = pagina.getBoundingClientRect().top;
    const vistos = new Set();
    editor.querySelectorAll("span.coment[data-c], ins.rev, del.rev").forEach((el) => {
      if (el.matches("span.coment")) {
        const id = el.dataset.c;
        if (vistos.has(id) || !comentarios[id]) return;
        vistos.add(id);
        itens.push({ chave: "c:" + id, tipo: "c", id, y: (el.getBoundingClientRect().top - base) / zoom, c: comentarios[id] });
      } else if (mostrarMarcas) {
        const anterior = itens[itens.length - 1];
        // Exclusão seguida de inclusão do mesmo autor = substituição (um balão só).
        if (anterior?.tipo === "r" && anterior.els.length === 1 && anterior.els[0].tagName === "DEL" && el.tagName === "INS" && el.previousSibling === anterior.els[0] && el.dataset.autor === anterior.els[0].dataset.autor) {
          anterior.els.push(el);
          return;
        }
        itens.push({ chave: "r:" + itens.length, tipo: "r", els: [el], y: (el.getBoundingClientRect().top - base) / zoom });
      }
    });
  }
  // Empilha os balões sem sobrepor, mantendo cada um o mais perto possível do seu trecho.
  useLayoutEffect(() => {
    const novos = {};
    let fim = 0;
    for (const it of itens) {
      const el = caixa.current?.querySelector(`[data-balao="${it.chave}"]`);
      const h = el?.offsetHeight || 60;
      const top = Math.max(it.y, fim);
      novos[it.chave] = top;
      fim = top + h + 6;
    }
    if (JSON.stringify(novos) !== JSON.stringify(tops)) setTops(novos);
  });
  const dataCurta = (iso) => { try { return new Date(iso).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }); } catch { return ""; } };
  const cartao = (cor, ativoIt) => ({ position: "absolute", left: 0, right: 0, background: "#fff", borderRadius: 8, padding: "8px 10px", fontSize: 12, fontFamily: "sans-serif", boxShadow: ativoIt ? "0 2px 12px rgba(0,0,0,0.25)" : "0 1px 4px rgba(0,0,0,0.15)", borderLeft: `4px solid ${cor}` });
  const bt = { border: "1px solid #ddd", background: "#fff", borderRadius: 4, padding: "2px 7px", fontSize: 11, cursor: "pointer" };
  const altura = Math.max(...itens.map((it) => (tops[it.chave] ?? it.y) + 140), 200);
  const irPara = (el) => el?.scrollIntoView({ block: "center", behavior: "smooth" });
  return (
    <aside ref={caixa} style={{ position: "relative", width: 260, flex: "none", minHeight: altura }}>
      {!itens.length && <div style={{ fontSize: 12, color: "#777", fontFamily: "sans-serif", padding: 8 }}>Comentários e alterações aparecem aqui, ao lado do trecho.</div>}
      {itens.map((it) => {
        const top = tops[it.chave] ?? it.y;
        if (it.tipo === "c") {
          const c = it.c;
          const ativoIt = ativo === it.id;
          return (
            <div key={it.chave} data-balao={it.chave} style={{ ...cartao(R.corAutor(c.autor), ativoIt), top, opacity: c.resolvido ? 0.6 : 1 }}
              onMouseEnter={() => setAtivo(it.id)} onMouseLeave={() => setAtivo(null)}
              onClick={() => irPara(editor.querySelector(`span.coment[data-c="${it.id}"]`))}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 6 }}>
                <b style={{ color: R.corAutor(c.autor) }}>💬 {c.autor}</b><span style={{ color: "#999" }}>{dataCurta(c.data)}</span>
              </div>
              {novo === it.id || !c.texto ? (
                <textarea autoFocus defaultValue={c.texto} placeholder="Escreva o comentário…" rows={3}
                  onClick={(e) => e.stopPropagation()}
                  onBlur={(e) => { mudar(it.id, { texto: e.target.value }); setNovo(null); }}
                  style={{ width: "100%", boxSizing: "border-box", marginTop: 4, fontSize: 12, fontFamily: "sans-serif", border: "1px solid #ddd", borderRadius: 4, padding: 4, resize: "vertical" }} />
              ) : (
                <div style={{ marginTop: 4, whiteSpace: "pre-wrap", color: "#222" }} onDoubleClick={(e) => { e.stopPropagation(); setNovo(it.id); }} title="Clique duplo para editar">{c.texto}</div>
              )}
              {(c.respostas || []).map((r, i) => (
                <div key={i} style={{ marginTop: 6, paddingTop: 4, borderTop: "1px solid #f0f0f0" }}>
                  <b style={{ color: R.corAutor(r.autor) }}>{r.autor}</b> <span style={{ color: "#999" }}>{dataCurta(r.data)}</span>
                  <div style={{ whiteSpace: "pre-wrap" }}>{r.texto}</div>
                </div>
              ))}
              {ativoIt && !c.resolvido && (
                <input placeholder="Responder… (Enter)" onClick={(e) => e.stopPropagation()}
                  onKeyDown={(e) => {
                    if (e.key !== "Enter" || !e.target.value.trim()) return;
                    mudar(it.id, { respostas: [...(c.respostas || []), { autor: autor || "Autor", data: new Date().toISOString(), texto: e.target.value.trim() }] });
                    e.target.value = "";
                  }}
                  style={{ width: "100%", boxSizing: "border-box", marginTop: 6, fontSize: 12, border: "1px solid #ddd", borderRadius: 4, padding: "3px 5px" }} />
              )}
              <div style={{ display: "flex", gap: 4, marginTop: 6 }} onClick={(e) => e.stopPropagation()}>
                <button style={bt} onClick={() => mudar(it.id, { resolvido: !c.resolvido })}>{c.resolvido ? "Reabrir" : "✓ Resolver"}</button>
                <button style={bt} onClick={() => excluir(it.id)}>🗑 Excluir</button>
              </div>
            </div>
          );
        }
        const [a, b] = it.els;
        const cor = R.corAutor(a.dataset.autor);
        const trecho = (el) => { const t = el.textContent.trim(); return t.length > 70 ? t.slice(0, 70) + "…" : t; };
        const desc = b ? <>Substituiu «<s>{trecho(a)}</s>» por «<u>{trecho(b)}</u>»</>
          : a.tagName === "INS" ? <>Incluiu «<u>{trecho(a)}</u>»</> : <>Excluiu «<s>{trecho(a)}</s>»</>;
        return (
          <div key={it.chave} data-balao={it.chave} style={{ ...cartao(cor), top, background: "#fbfbfb" }} onClick={() => irPara(a)}>
            <div style={{ display: "flex", justifyContent: "space-between", gap: 6 }}>
              <b style={{ color: cor }}>✎ {a.dataset.autor}</b><span style={{ color: "#999" }}>{dataCurta(a.dataset.data)}</span>
            </div>
            <div style={{ marginTop: 3, color: "#333", wordBreak: "break-word" }}>{desc}</div>
            <div style={{ display: "flex", gap: 4, marginTop: 5 }} onClick={(e) => e.stopPropagation()}>
              <button style={bt} onClick={() => it.els.forEach((el) => revisar(el, true))}>✓ Aceitar</button>
              <button style={bt} onClick={() => it.els.forEach((el) => revisar(el, false))}>✕ Rejeitar</button>
            </div>
          </div>
        );
      })}
    </aside>
  );
}
