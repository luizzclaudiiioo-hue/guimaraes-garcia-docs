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
const CSS_FONTES = ["Arial Narrow", "Arial Nova Cond"].flatMap((nome) => [[fonte400, 400, "normal"], [fonte400i, 400, "italic"], [fonte700, 700, "normal"], [fonte700i, 700, "italic"]]
  .map(([url, peso, estilo]) => `@font-face { font-family: "${nome}"; font-weight: ${peso}; font-style: ${estilo}; src: local("${nome}"), local("Arial Narrow"), url("${url}") format("woff2"); }`)).join("\n");

const sBarra = { display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center", padding: 8, background: "#f4f1ea", border: "1px solid rgba(0,0,0,0.12)", borderRadius: "10px 10px 0 0", position: "sticky", top: 0, zIndex: 5 };
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
.ed-texto img { max-width: 100%; }
.ed-area { font-size: 10pt; line-height: 1.3; min-height: 1.2em; outline: none; border: 1px dashed transparent; border-radius: 4px; padding: 2px 4px; }
.ed-area:hover, .ed-area:focus { border-color: #d8cfb8; }
.ed-area:empty::before { content: attr(data-vazio); color: #b9b2a3; }
.ed-area p { margin: 0; }
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
  const [painelPagina, setPainelPagina] = useState(false);
  const [menu, setMenu] = useState(null); // menu do botão direito: { x, y, td }
  const [seletorTabela, setSeletorTabela] = useState(null); // { l, c } sob o mouse
  const selCelulas = useRef(null); // { tabela, inicio, ret }
  const [imgSel, setImgSel] = useState(null); // imagem selecionada (elemento)
  const [painelEstilos, setPainelEstilos] = useState(false);
  const [autor, setAutor] = useState(() => ler("editorTexto.autor", ""));
  const [mostrarMarcas, setMostrarMarcas] = useState(true);
  const [comAtivo, setComAtivo] = useState(null); // comentário destacado
  const [comNovo, setComNovo] = useState(null); // comentário recém-criado (abre para escrever)
  const [menuRevisao, setMenuRevisao] = useState(false);
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
        insertUnorderedList: document.queryCommandState("insertUnorderedList"),
        insertOrderedList: document.queryCommandState("insertOrderedList"),
        bloco: (document.queryCommandValue("formatBlock") || "p").toLowerCase(),
      });
      // Página onde está o cursor
      if (s.rangeCount) {
        const r = s.getRangeAt(0).getBoundingClientRect();
        const topo = editor.current.getBoundingClientRect().top;
        const y = (r.top || r.bottom) - topo;
        if (r.top || r.bottom) setPagAtual(quebrasRef.current.filter((q) => q <= y).length + 1);
      }
    };
    document.addEventListener("selectionchange", atualizarEstado);
    return () => document.removeEventListener("selectionchange", atualizarEstado);
  }, []);

  // ----- páginas -----

  const quebrasRef = useRef([]);
  const medindo = useRef(false);
  function medirPaginas() {
    if (medindo.current) return;
    medindo.current = true;
    requestAnimationFrame(() => {
      medindo.current = false;
      const ed = editor.current;
      if (!ed) return;
      const H = alturaUtilRef.current;
      const lista = [];
      let inicio = 0;
      for (const filho of ed.children) {
        const topo = filho.offsetTop, fim = topo + filho.offsetHeight;
        if (filho.tagName === "HR") { lista.push(fim); inicio = fim; continue; }
        while (fim > inicio + H + 1) {
          // Parágrafo curto que não cabe vai inteiro para a próxima página; longo é dividido.
          const quebra = topo > inicio && fim - topo < 48 ? topo : inicio + H;
          lista.push(quebra);
          inicio = quebra;
        }
      }
      quebrasRef.current = lista;
      atualizarSumarios(lista);
      setQuebras((a) => (a.length === lista.length && a.every((v, i) => Math.abs(v - lista[i]) < 1) ? a : lista));
    });
  }
  const alturaUtilRef = useRef(alturaUtil);
  alturaUtilRef.current = alturaUtil;
  useLayoutEffect(() => { medirPaginas(); }, [alturaUtil, largCm, m.esq, m.dir]);

  // ----- alterações -----

  function mudarDoc(patch, marcarAlterado = true) {
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, ...patch, atualizado: Date.now(), ...(marcarAlterado ? { alterado: true } : {}) } : x)));
  }

  // HTML do editor sem marcas temporárias da tela (células selecionadas).
  const htmlLimpo = () => editor.current.innerHTML.replace(/ data-sel="1"/g, "");

  function salvarHtml() {
    renumerarNotas();
    mudarDoc({ html: htmlLimpo() });
    medirPaginas();
  }

  function cmd(c, v = null) {
    editor.current.focus();
    document.execCommand(c, false, v);
    salvarHtml();
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
    setMenuRevisao(false);
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
    const w0 = img.getBoundingClientRect().width, h0 = img.getBoundingClientRect().height;
    const proporcao = h0 / w0;
    const max = editor.current.getBoundingClientRect().width;
    const x0 = e.clientX;
    const sinal = canto.includes("d") ? 1 : -1; // alças da direita crescem para a direita
    const mover = (ev) => {
      const w = Math.max(24, Math.min(max, w0 + sinal * (ev.clientX - x0)));
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
    const max = editor.current.getBoundingClientRect().width;
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
    if (imgSel && (e.key === "Delete" || e.key === "Backspace")) { e.preventDefault(); return excluirImagem(); }
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
        novo = criar(nome, { html: sanitizar(conv.html || "<p><br></p>"), cabecalho: sanitizar(conv.cabecalho || ""), rodape: sanitizar(conv.rodape || ""), pagina: conv.pagina || PAGINA_PADRAO, comentarios: conv.comentarios || {}, notas: Object.fromEntries(Object.entries(conv.notas || {}).map(([k, v]) => [k, sanitizar(v)])), ...vinculo, avisos });
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
    const conteudo = { html: htmlLimpo(), cabecalho: cab.current.innerHTML, rodape: rod.current.innerHTML, notas: atual.notas || {}, comentarios: atual.comentarios || {} };
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

  function imprimir() {
    const w = window.open("", "_blank");
    if (!w) return setMsg("Permita pop-ups para imprimir ou salvar em PDF.");
    const numero = {
      direita: `@bottom-right { content: counter(page); font: 10pt "${estilos.normal.fonte}"; }`,
      centro: `@bottom-center { content: counter(page); font: 10pt "${estilos.normal.fonte}"; }`,
      xdey: `@bottom-center { content: "Página " counter(page) " de " counter(pages); font: 10pt "${estilos.normal.fonte}"; }`,
    }[atual.numPag] || "";
    // Cabeçalho/rodapé em thead/tfoot: o navegador repete em todas as páginas impressas.
    w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(atual.nome)}</title>
      <style>
        @page { size: ${largCm}cm ${altCm}cm; margin: ${m.sup}cm ${m.dir}cm ${m.inf}cm ${m.esq}cm; ${numero} }
        ${CSS_FONTES.replace(/url\("\//g, `url("${location.origin}/`)}
        body { margin: 0; } ${CSS_PAGINA} ${cssDoc}
        .sumario { border: 0; padding: 0; } .ed-notas { margin-top: 18pt; border-top: 1px solid #000; padding-top: 4pt; }
        table.moldura { width: 100%; border-collapse: collapse; } table.moldura > * > tr > td { padding: 0; }
        .ed-area { border: 0; padding: 0; } .ed-area:empty { display: none; }
        hr { page-break-after: always; break-after: page; border: 0; margin: 0; }
      </style></head><body>
      <table class="moldura">
        ${cab.current.innerHTML ? `<thead><tr><td><div class="ed-area" style="padding-bottom:12pt">${cab.current.innerHTML}</div></td></tr></thead>` : ""}
        ${rod.current.innerHTML ? `<tfoot><tr><td><div class="ed-area" style="padding-top:12pt">${rod.current.innerHTML}</div></td></tr></tfoot>` : ""}
        <tbody><tr><td><div class="ed-texto">${htmlLimpo()}</div>${ordemNotas.length ? `<div class="ed-notas">${ordemNotas.map((id, i) => `<div style="display:flex;gap:6px"><sup>${i + 1}</sup><div>${atual.notas?.[id] || ""}</div></div>`).join("")}</div>` : ""}</td></tr></tbody>
      </table></body></html>`);
    w.document.close();
    w.focus();
    w.print();
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
  function localizar(paraTras = false) {
    const { termo, caso } = busca;
    const { lista, nos } = ocorrencias(termo, caso);
    if (!lista.length) return setBusca((b) => ({ ...b, info: "Nada encontrado" }));
    const s = document.getSelection();
    let pos = posicaoCursor(nos);
    if (paraTras && s.rangeCount) pos -= s.toString().length;
    let i = paraTras ? [...lista].reverse().findIndex((x) => x < pos) : lista.findIndex((x) => x >= pos);
    if (paraTras) i = i < 0 ? lista.length - 1 : lista.length - 1 - i;
    else if (i < 0) i = 0;
    selecionar(intervalo(nos, lista[i], lista[i] + termo.length));
    setBusca((b) => ({ ...b, info: `${i + 1} de ${lista.length}` }));
  }
  function iguais(a, b, caso) {
    return caso ? a === b : a.toLocaleLowerCase("pt-BR") === b.toLocaleLowerCase("pt-BR");
  }
  function substituir() {
    const { termo, novo, caso } = busca;
    const s = document.getSelection();
    if (s.rangeCount && editor.current.contains(s.anchorNode) && iguais(s.toString(), termo, caso)) {
      document.execCommand("insertText", false, novo);
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
    setBusca((b) => ({ ...b, info: lista.length ? `${lista.length} substituição(ões) feitas` : "Nada encontrado" }));
  }
  function abrirBusca(substituirTambem) {
    const s = document.getSelection();
    const sel = s.rangeCount && editor.current.contains(s.anchorNode) ? s.toString() : "";
    setBusca((b) => ({ termo: sel || b?.termo || "", novo: b?.novo || "", caso: b?.caso || false, info: "", subst: substituirTambem || b?.subst }));
    setTimeout(() => campoBusca.current?.focus(), 0);
  }

  function teclas(e) {
    if (editor.current?.contains(e.target)) teclaTabela(e);
    if (e.defaultPrevented || !(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    if (k === "s") { e.preventDefault(); salvarDocx(e.shiftKey); }
    if (k === "p") { e.preventDefault(); imprimir(); }
    if (k === "l" || k === "f") { e.preventDefault(); abrirBusca(false); }
    if (k === "h") { e.preventDefault(); abrirBusca(true); }
    if (k === "z" && !e.shiftKey && atual.controlar && editor.current?.contains(e.target)) { e.preventDefault(); desfazerControle(); }
  }

  const arquivoAtual = atual.arquivo && (arquivos.current.get(atual.id)?.name || atual.arquivo);
  const palavras = (atual.html.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").match(/\S+/g) || []).length;
  const totalPag = quebras.length + 1;
  const textoNumero = { direita: "1", centro: "1", xdey: `Página 1 de ${totalPag}` }[atual.numPag];

  return (
    <div style={{ fontFamily: "sans-serif" }} onKeyDown={teclas}>
      <style>{CSS_FONTES + CSS_PAGINA + cssDoc + CSS_TELA + cssRevisao}</style>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
        <select value={atualId} onChange={(e) => trocar(e.target.value)} style={{ ...sSelect, maxWidth: 220 }} title="Documentos salvos neste navegador">
          {docs.map((x) => <option key={x.id} value={x.id}>{x.nome || "(sem nome)"}</option>)}
        </select>
        <input value={atual.nome} onChange={(e) => renomear(e.target.value)} style={{ ...sSelect, flex: "1 1 180px", padding: "0 8px" }} placeholder="Nome do documento" />
        <button style={sBotao()} onClick={() => criar()}>＋ Novo</button>
        <button style={sBotao()} onClick={abrir} title="Abrir .docx, .txt ou .html do computador">📂 Abrir</button>
        <button style={sBotao()} onClick={() => salvarDocx(false)} title={arquivoAtual ? `Gravar no arquivo "${arquivoAtual}" (Ctrl+S)` : "Salvar em Word (.docx) (Ctrl+S)"}>💾 Salvar</button>
        <button style={sBotao()} onClick={() => salvarDocx(true)} title="Salvar uma cópia em outro arquivo (Ctrl+Shift+S)">Salvar como…</button>
        <button style={sBotao()} onClick={imprimir}>🖨️ PDF / Imprimir</button>
        <button style={sBotao()} onClick={excluir} title="Excluir documento">🗑️</button>
        <input ref={arquivo} type="file" accept=".docx,.txt,.md,.html,.htm" onChange={abrirArquivo} style={{ display: "none" }} />
        <input ref={arquivoImg} type="file" accept="image/png,image/jpeg,image/gif,image/bmp" onChange={inserirImagem} style={{ display: "none" }} />
      </div>
      <p style={{ fontSize: 12, color: "#666", margin: "-4px 0 8px" }}>
        {arquivoAtual
          ? <>📄 Arquivo: <b>{arquivoAtual}</b> · {atual.alterado ? <span style={{ color: "#b0471b" }}>● alterações não salvas (Ctrl+S grava no arquivo)</span> : <span style={{ color: "#4a7a4a" }}>✓ salvo</span>}</>
          : suportaArquivoLocal ? "Documento ainda não salvo em arquivo: \"Salvar\" pergunta onde guardar e, depois disso, grava sempre no mesmo arquivo."
          : "Este navegador não permite salvar no mesmo arquivo: \"Salvar\" baixa um .docx novo. Para gravar no próprio arquivo, use Chrome ou Edge no computador."}
      </p>
      {msg && <p style={{ fontSize: 13, color: GOLD, margin: "0 0 8px" }}>{msg}</p>}

      <div style={sBarra} onMouseDown={(e) => { if (e.target.tagName === "BUTTON") e.preventDefault(); }}>
        <button style={sBotao()} onClick={() => cmd("undo")} title="Desfazer (Ctrl+Z)">↶</button>
        <button style={sBotao()} onClick={() => cmd("redo")} title="Refazer (Ctrl+Y)">↷</button>
        <span style={sSep} />
        <select style={sSelect} value={["h1", "h2", "h3", "blockquote"].includes(estado.bloco) ? estado.bloco : "p"} onChange={(e) => cmd("formatBlock", e.target.value)} title="Estilo">
          <option value="p">Normal</option>
          <option value="h1">Título 1</option>
          <option value="h2">Título 2</option>
          <option value="h3">Título 3</option>
          <option value="blockquote">Citação</option>
        </select>
        <select style={{ ...sSelect, width: 120 }} defaultValue="" onChange={(e) => { cmd("fontName", e.target.value); e.target.value = ""; }} title="Fonte">
          <option value="" disabled>Fonte</option>
          {FONTES.map((f) => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
        </select>
        <select style={sSelect} defaultValue="" onChange={(e) => { aplicarTamanho(e.target.value); e.target.value = ""; }} title="Tamanho da fonte (pt)">
          <option value="" disabled>Tam.</option>
          {TAMANHOS.map((pt) => <option key={pt} value={pt}>{String(pt).replace(".", ",")}</option>)}
        </select>
        <span style={sSep} />
        <button style={{ ...sBotao(estado.bold), fontWeight: "bold" }} onClick={() => cmd("bold")} title="Negrito (Ctrl+B)">N</button>
        <button style={{ ...sBotao(estado.italic), fontStyle: "italic" }} onClick={() => cmd("italic")} title="Itálico (Ctrl+I)">I</button>
        <button style={{ ...sBotao(estado.underline), textDecoration: "underline" }} onClick={() => cmd("underline")} title="Sublinhado (Ctrl+U)">S</button>
        <button style={{ ...sBotao(estado.strikeThrough), textDecoration: "line-through" }} onClick={() => cmd("strikeThrough")} title="Tachado">ab</button>
        <label style={{ ...sBotao(), display: "inline-flex", alignItems: "center", gap: 2 }} title="Cor do texto">
          A<input type="color" onChange={(e) => cmd("foreColor", e.target.value)} style={{ width: 18, height: 18, border: 0, padding: 0, background: "none" }} />
        </label>
        <label style={{ ...sBotao(), display: "inline-flex", alignItems: "center", gap: 2 }} title="Realce">
          🖍<input type="color" defaultValue="#ffff00" onChange={(e) => cmd("hiliteColor", e.target.value)} style={{ width: 18, height: 18, border: 0, padding: 0, background: "none" }} />
        </label>
        <span style={sSep} />
        <button style={sBotao()} onClick={() => cmd("justifyLeft")} title="Alinhar à esquerda">⯇</button>
        <button style={sBotao()} onClick={() => cmd("justifyCenter")} title="Centralizar">≡</button>
        <button style={sBotao()} onClick={() => cmd("justifyRight")} title="Alinhar à direita">⯈</button>
        <button style={sBotao()} onClick={() => cmd("justifyFull")} title="Justificar">☰</button>
        <span style={sSep} />
        <select style={{ ...sSelect, width: 112 }} value="" onChange={(e) => estiloParagrafo("lineHeight", e.target.value)} title="Espaçamento entre linhas">
          <option value="" disabled>↕ Entrelinhas</option>
          {["1", "1.15", "1.5", "2"].map((v) => <option key={v} value={v}>{v.replace(".", ",")}</option>)}
        </select>
        <select style={{ ...sSelect, width: 128 }} value="" onChange={(e) => estiloParagrafo("marginBottom", e.target.value)} title="Espaço depois do parágrafo">
          <option value="" disabled>¶ Espaço depois</option>
          {["0pt", "6pt", "12pt", "18pt"].map((v) => <option key={v} value={v}>{v.replace("pt", " pt")}</option>)}
        </select>
        <select style={{ ...sSelect, width: 118 }} value="" onChange={(e) => estiloParagrafo("textIndent", e.target.value)} title="Recuo da primeira linha">
          <option value="" disabled>⇥ Recuo 1ª linha</option>
          <option value="0cm">Nenhum</option>
          <option value="1.25cm">1,25 cm</option>
          <option value="2.5cm">2,5 cm</option>
          <option value="4cm">4 cm (citação)</option>
        </select>
        <span style={sSep} />
        <button style={sBotao(estado.insertUnorderedList)} onClick={() => cmd("insertUnorderedList")} title="Marcadores">• Lista</button>
        <button style={sBotao(estado.insertOrderedList)} onClick={() => cmd("insertOrderedList")} title="Numeração">1. Lista</button>
        <button style={sBotao()} onClick={() => cmd("outdent")} title="Diminuir recuo">⇤</button>
        <button style={sBotao()} onClick={() => cmd("indent")} title="Aumentar recuo">⇥</button>
        <span style={sSep} />
        <span style={{ position: "relative" }}>
          <button style={sBotao(!!seletorTabela)} onClick={() => setSeletorTabela((x) => (x ? null : { l: 0, c: 0 }))} title="Inserir tabela">▦ Tabela</button>
          {seletorTabela && (
            <div style={{ position: "absolute", top: 34, left: 0, zIndex: 20, background: "#fff", border: "1px solid #ccc", borderRadius: 6, padding: 8, boxShadow: "0 4px 16px rgba(0,0,0,0.15)" }}
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
        <button style={sBotao()} onClick={() => arquivoImg.current.click()} title="Inserir imagem">🖼 Imagem</button>
        <button style={sBotao()} onClick={inserirLink} title="Inserir link">🔗</button>
        <button style={sBotao()} onClick={() => cmd("insertHTML", "<hr><p><br></p>")} title="Quebra de página">⤓ Quebra</button>
        <button style={sBotao()} onClick={inserirNota} title="Inserir nota de rodapé no ponto do cursor">¹ Nota</button>
        <button style={sBotao()} onClick={inserirSumario} title="Inserir sumário automático (usa os Títulos 1, 2 e 3)">≣ Sumário</button>
        <button style={sBotao(painelEstilos)} onClick={() => setPainelEstilos((x) => !x)} title="Editar os estilos Normal, Títulos e Citação">Aa Estilos</button>
        <span style={sSep} />
        <button style={sBotao()} onClick={comentar} title="Comentar o trecho selecionado">💬 Comentar</button>
        <button style={{ ...sBotao(atual.controlar), ...(atual.controlar ? { border: "1px solid #b0471b", color: "#b0471b" } : {}) }} onClick={alternarControle}
          title="Controlar alterações: inclusões e exclusões ficam marcadas com o seu nome">✎ Controlar alterações{atual.controlar ? ": ligado" : ""}</button>
        <span style={{ position: "relative" }}>
          <button style={sBotao(menuRevisao)} onClick={() => setMenuRevisao((x) => !x)} title="Aceitar ou rejeitar alterações e mostrar/ocultar marcações">Revisão ▾</button>
          {menuRevisao && (
            <div style={{ position: "absolute", top: 34, right: 0, zIndex: 20, background: "#fff", border: "1px solid #ccc", borderRadius: 6, padding: 6, boxShadow: "0 4px 16px rgba(0,0,0,0.15)", display: "flex", flexDirection: "column", gap: 4, minWidth: 230 }}>
              <button style={sBotao()} onClick={() => revisarTodas(true)}>✓ Aceitar todas as alterações</button>
              <button style={sBotao()} onClick={() => revisarTodas(false)}>✕ Rejeitar todas as alterações</button>
              <label style={{ fontSize: 13, display: "flex", gap: 6, alignItems: "center", padding: "4px 2px" }}>
                <input type="checkbox" checked={mostrarMarcas} onChange={(e) => setMostrarMarcas(e.target.checked)} /> Mostrar marcações
              </label>
              <button style={sBotao()} onClick={() => { setMenuRevisao(false); const n = prompt("Seu nome (aparece nos comentários e nas alterações):", autor)?.trim(); if (n) { gravar("editorTexto.autor", n); setAutor(n); } }}>Meu nome: {autor || "(não definido)"}</button>
            </div>
          )}
        </span>
        <button style={sBotao()} onClick={() => cmd("removeFormat")} title="Limpar formatação">⌫ Formato</button>
        <span style={sSep} />
        <button style={sBotao(!!busca)} onClick={() => (busca ? setBusca(null) : abrirBusca(true))} title="Localizar e substituir (Ctrl+L / Ctrl+H)">🔍 Localizar</button>
        <button style={sBotao(painelPagina)} onClick={() => setPainelPagina((x) => !x)} title="Tamanho do papel, margens, cabeçalho, rodapé e número de página">📄 Página</button>
      </div>

      {busca && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center", padding: 8, background: "#fbf8f1", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0 }}>
          <input ref={campoBusca} placeholder="Localizar" value={busca.termo} onChange={(e) => setBusca({ ...busca, termo: e.target.value, info: "" })}
            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); localizar(e.shiftKey); } if (e.key === "Escape") setBusca(null); }}
            style={{ ...sCampo, width: 200 }} />
          <input placeholder="Substituir por" value={busca.novo} onChange={(e) => setBusca({ ...busca, novo: e.target.value })}
            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); substituir(); } if (e.key === "Escape") setBusca(null); }}
            style={{ ...sCampo, width: 200 }} />
          <label style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 4 }}>
            <input type="checkbox" checked={busca.caso} onChange={(e) => setBusca({ ...busca, caso: e.target.checked, info: "" })} /> Diferenciar maiúsculas
          </label>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={() => localizar(true)} disabled={!busca.termo} title="Anterior (Shift+Enter)">↑</button>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={() => localizar(false)} disabled={!busca.termo} title="Próximo (Enter)">↓ Próximo</button>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={substituir} disabled={!busca.termo}>Substituir</button>
          <button style={sBotao()} onMouseDown={(e) => e.preventDefault()} onClick={substituirTudo} disabled={!busca.termo}>Substituir tudo</button>
          <span style={{ fontSize: 12, color: "#666" }}>{busca.info}</span>
          <button style={{ ...sBotao(), marginLeft: "auto" }} onClick={() => setBusca(null)} title="Fechar (Esc)">×</button>
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

      <div style={{ background: "#e9e6df", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0, borderRadius: "0 0 10px 10px", padding: "24px 12px", overflowX: "auto", display: "flex", justifyContent: "center", alignItems: "flex-start", gap: 14 }}>
        <div ref={pagRef} style={{ position: "relative", background: "#fff", width: `${largCm}cm`, maxWidth: "100%", flex: "none", minHeight: `${altCm}cm`, boxShadow: "0 2px 12px rgba(0,0,0,0.15)", boxSizing: "border-box", padding: `0 ${m.dir}cm 0 ${m.esq}cm`, display: "flex", flexDirection: "column" }}>
          <div style={{ minHeight: `${m.sup}cm`, display: "flex", alignItems: "flex-end", paddingBottom: "0.3cm", boxSizing: "border-box" }}>
            <div ref={cab} className="ed-area" contentEditable suppressContentEditableWarning data-vazio="Cabeçalho (clique para escrever)" lang="pt-BR" spellCheck
              onInput={() => mudarDoc({ cabecalho: cab.current.innerHTML })} style={{ width: "100%", fontFamily: `'${estilos.normal.fonte}', sans-serif` }} />
          </div>
          <div style={{ position: "relative", flex: 1 }}>
            <div ref={editor} className={"ed-texto" + (mostrarMarcas ? "" : " sem-marcas")} contentEditable suppressContentEditableWarning spellCheck lang="pt-BR"
              onInput={salvarHtml} onClick={cliqueEditor} onMouseDown={mouseDownEditor} onMouseMove={mouseMoveEditor} onContextMenu={menuContexto}
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
            {imgSel?.isConnected && <MolduraImagem img={imgSel} base={editor.current} redimensionar={redimensionarImagem} posicionar={posicionarImagem} original={tamanhoOriginal} excluir={excluirImagem} />}
            {quebras.map((y, i) => (
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
        {(Object.keys(comentarios).length > 0 || autoresRev.length > 0 || atual.controlar) && (
          <PainelRevisao editor={editor.current} pagina={pagRef.current} comentarios={comentarios} versao={atual.html} quebras={quebras}
            ativo={comAtivo} setAtivo={setComAtivo} novo={comNovo} setNovo={setComNovo} autor={autor} mostrarMarcas={mostrarMarcas}
            mudar={mudarComentario} excluir={excluirComentario} revisar={revisar} />
        )}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8, fontSize: 12, color: "#666", marginTop: 6 }}>
        <span>Página {Math.min(pagAtual, totalPag)} de {totalPag} · {palavras} palavra{palavras === 1 ? "" : "s"} · salvo automaticamente neste navegador</span>
        <span>Ctrl+S salva · Ctrl+Shift+S salva como · Ctrl+P imprime · Ctrl+L localiza · Ctrl+H substitui</span>
      </div>
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

function MolduraImagem({ img, base, redimensionar, posicionar, original, excluir }) {
  const b = base.getBoundingClientRect(), r = img.getBoundingClientRect();
  const x = r.left - b.left, y = r.top - b.top;
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
function PainelRevisao({ editor, pagina, comentarios, versao, quebras, ativo, setAtivo, novo, setNovo, autor, mostrarMarcas, mudar, excluir, revisar }) {
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
        itens.push({ chave: "c:" + id, tipo: "c", id, y: el.getBoundingClientRect().top - base, c: comentarios[id] });
      } else if (mostrarMarcas) {
        const anterior = itens[itens.length - 1];
        // Exclusão seguida de inclusão do mesmo autor = substituição (um balão só).
        if (anterior?.tipo === "r" && anterior.els.length === 1 && anterior.els[0].tagName === "DEL" && el.tagName === "INS" && el.previousSibling === anterior.els[0] && el.dataset.autor === anterior.els[0].dataset.autor) {
          anterior.els.push(el);
          return;
        }
        itens.push({ chave: "r:" + itens.length, tipo: "r", els: [el], y: el.getBoundingClientRect().top - base });
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
