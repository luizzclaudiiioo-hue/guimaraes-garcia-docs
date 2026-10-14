import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { CM, PAGINA_PADRAO, dimensoes, docxParaHtml, htmlParaDocx } from "./docxConversao.js";
import * as T from "./textoTabelas.js";
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

const FONTES = ["Times New Roman", "Arial", "Calibri", "Georgia", "Verdana", "Courier New"];
const TAMANHOS = [["1", "8"], ["2", "10"], ["3", "12"], ["4", "14"], ["5", "18"], ["6", "24"], ["7", "36"]];

const sBarra = { display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center", padding: 8, background: "#f4f1ea", border: "1px solid rgba(0,0,0,0.12)", borderRadius: "10px 10px 0 0", position: "sticky", top: 0, zIndex: 5 };
const sBotao = (ativo) => ({ minWidth: 32, height: 30, padding: "0 8px", border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: ativo ? "rgba(154,110,30,0.2)" : "#fff", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif", color: "#2a2a2a", whiteSpace: "nowrap" });
const sSelect = { height: 30, border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: "#fff", fontSize: 13, fontFamily: "sans-serif", padding: "0 4px" };
const sSep = { width: 1, height: 22, background: "rgba(0,0,0,0.15)", margin: "0 4px" };
const sCampo = { ...sSelect, padding: "0 8px" };

// Aparência do texto na página; o .docx gerado usa os mesmos padrões.
const CSS_PAGINA = `
.ed-texto { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.5; color: #111; outline: none; }
.ed-texto p, .ed-texto div, .ed-texto li { margin: 0 0 6pt; }
.ed-texto h1 { font-size: 16pt; margin: 12pt 0 6pt; } .ed-texto h2 { font-size: 14pt; margin: 12pt 0 6pt; } .ed-texto h3 { font-size: 12pt; margin: 12pt 0 6pt; }
.ed-texto blockquote { margin: 0 0 6pt 1.27cm; }
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
    mudarDoc({ html: htmlLimpo() });
    medirPaginas();
  }

  function cmd(c, v = null) {
    editor.current.focus();
    document.execCommand(c, false, v);
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
        novo = criar(nome, { html: sanitizar(conv.html || "<p><br></p>"), cabecalho: sanitizar(conv.cabecalho || ""), rodape: sanitizar(conv.rodape || ""), pagina: conv.pagina || PAGINA_PADRAO, ...vinculo, avisos });
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
    const conteudo = { html: htmlLimpo(), cabecalho: cab.current.innerHTML, rodape: rod.current.innerHTML };
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
      direita: '@bottom-right { content: counter(page); font: 10pt "Times New Roman"; }',
      centro: '@bottom-center { content: counter(page); font: 10pt "Times New Roman"; }',
      xdey: '@bottom-center { content: "Página " counter(page) " de " counter(pages); font: 10pt "Times New Roman"; }',
    }[atual.numPag] || "";
    // Cabeçalho/rodapé em thead/tfoot: o navegador repete em todas as páginas impressas.
    w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(atual.nome)}</title>
      <style>
        @page { size: ${largCm}cm ${altCm}cm; margin: ${m.sup}cm ${m.dir}cm ${m.inf}cm ${m.esq}cm; ${numero} }
        body { margin: 0; } ${CSS_PAGINA}
        table.moldura { width: 100%; border-collapse: collapse; } table.moldura > * > tr > td { padding: 0; }
        .ed-area { border: 0; padding: 0; } .ed-area:empty { display: none; }
        hr { page-break-after: always; break-after: page; border: 0; margin: 0; }
      </style></head><body>
      <table class="moldura">
        ${cab.current.innerHTML ? `<thead><tr><td><div class="ed-area" style="padding-bottom:12pt">${cab.current.innerHTML}</div></td></tr></thead>` : ""}
        ${rod.current.innerHTML ? `<tfoot><tr><td><div class="ed-area" style="padding-top:12pt">${rod.current.innerHTML}</div></td></tr></tfoot>` : ""}
        <tbody><tr><td><div class="ed-texto">${editor.current.innerHTML}</div></td></tr></tbody>
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
  }

  const arquivoAtual = atual.arquivo && (arquivos.current.get(atual.id)?.name || atual.arquivo);
  const palavras = (atual.html.replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").match(/\S+/g) || []).length;
  const totalPag = quebras.length + 1;
  const textoNumero = { direita: "1", centro: "1", xdey: `Página 1 de ${totalPag}` }[atual.numPag];

  return (
    <div style={{ fontFamily: "sans-serif" }} onKeyDown={teclas}>
      <style>{CSS_PAGINA + CSS_TELA}</style>
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
        <select style={sSelect} defaultValue="" onChange={(e) => { cmd("fontSize", e.target.value); e.target.value = ""; }} title="Tamanho">
          <option value="" disabled>Tam.</option>
          {TAMANHOS.map(([v, pt]) => <option key={v} value={v}>{pt}</option>)}
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

      <div style={{ background: "#e9e6df", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0, borderRadius: "0 0 10px 10px", padding: "24px 12px", overflowX: "auto" }}>
        <div style={{ position: "relative", background: "#fff", width: `${largCm}cm`, maxWidth: "100%", minHeight: `${altCm}cm`, margin: "0 auto", boxShadow: "0 2px 12px rgba(0,0,0,0.15)", boxSizing: "border-box", padding: `0 ${m.dir}cm 0 ${m.esq}cm`, display: "flex", flexDirection: "column" }}>
          <div style={{ minHeight: `${m.sup}cm`, display: "flex", alignItems: "flex-end", paddingBottom: "0.3cm", boxSizing: "border-box" }}>
            <div ref={cab} className="ed-area" contentEditable suppressContentEditableWarning data-vazio="Cabeçalho (clique para escrever)" lang="pt-BR" spellCheck
              onInput={() => mudarDoc({ cabecalho: cab.current.innerHTML })} style={{ width: "100%", fontFamily: "'Times New Roman', serif" }} />
          </div>
          <div style={{ position: "relative", flex: 1 }}>
            <div ref={editor} className="ed-texto" contentEditable suppressContentEditableWarning spellCheck lang="pt-BR"
              onInput={salvarHtml} onMouseDown={mouseDownEditor} onMouseMove={mouseMoveEditor} onContextMenu={menuContexto}
              style={{ minHeight: alturaUtil, position: "relative" }} />
            {quebras.map((y, i) => (
              <div key={i} style={{ position: "absolute", left: `-${m.esq}cm`, right: `-${m.dir}cm`, top: y, borderTop: "1px dashed #c9bfa8", pointerEvents: "none" }}>
                <span style={{ position: "absolute", right: 4, top: -9, fontSize: 10, color: "#a8987a", background: "#fff", padding: "0 4px", fontFamily: "sans-serif" }}>Página {i + 2}</span>
              </div>
            ))}
          </div>
          <div style={{ minHeight: `${m.inf}cm`, paddingTop: "0.3cm", boxSizing: "border-box" }}>
            <div ref={rod} className="ed-area" contentEditable suppressContentEditableWarning data-vazio="Rodapé (clique para escrever)" lang="pt-BR" spellCheck
              onInput={() => mudarDoc({ rodape: rod.current.innerHTML })} style={{ width: "100%", fontFamily: "'Times New Roman', serif" }} />
            {textoNumero && (
              <div style={{ fontSize: "10pt", fontFamily: "'Times New Roman', serif", textAlign: atual.numPag === "direita" ? "right" : "center", color: "#555", padding: "0 4px" }} title="Número de página (aparece em todas as páginas no Word e na impressão)">
                {textoNumero}
              </div>
            )}
          </div>
        </div>
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
