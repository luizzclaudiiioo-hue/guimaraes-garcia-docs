// Controle de alterações (como no Word) e âncoras de comentários no editor de Texto.
//
// Marcação usada no HTML do editor:
//   <ins class="rev" data-autor="…" data-data="ISO">texto incluído</ins>
//   <del class="rev" data-autor="…" data-data="ISO">texto excluído</del>
//   <span class="coment" data-c="id">trecho comentado</span>  (pode repetir o mesmo id)

const agora = () => new Date().toISOString();

// Nós de texto dentro do intervalo, já cortados nas bordas.
export function textosNoIntervalo(range) {
  const raiz = range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer;
  const w = document.createTreeWalker(raiz, NodeFilter.SHOW_TEXT);
  const lista = [];
  for (let n = w.nextNode(); n; n = w.nextNode()) if (range.intersectsNode(n)) lista.push(n);
  const out = [];
  for (const t of lista) {
    const ini = t === range.startContainer ? range.startOffset : 0;
    const fim = t === range.endContainer ? range.endOffset : t.length;
    if (ini >= fim) continue;
    let seg = t;
    if (ini > 0) seg = seg.splitText(ini);
    if (fim - ini < seg.length) seg.splitText(fim - ini);
    out.push(seg);
  }
  return out;
}

// Envolve cada pedaço de texto do intervalo num elemento criado por `criar()`.
export function envolverIntervalo(range, criar) {
  const segs = textosNoIntervalo(range);
  const criados = [];
  for (const seg of segs) {
    if (!seg.data) continue;
    const el = criar();
    seg.parentNode.insertBefore(el, seg);
    el.appendChild(seg);
    criados.push(el);
  }
  return criados;
}

export function desembrulhar(el) {
  el.replaceWith(...el.childNodes);
}

function juntarVizinhos(el) {
  // Junta <del>/<ins> do mesmo autor encostados um no outro.
  for (const lado of ["previousSibling", "nextSibling"]) {
    const v = el[lado];
    if (v && v.nodeType === 1 && v.tagName === el.tagName && v.classList.contains("rev") && v.dataset.autor === el.dataset.autor) {
      if (lado === "previousSibling") { v.append(...el.childNodes); el.remove(); el = v; }
      else { el.append(...v.childNodes); v.remove(); }
    }
  }
  return el;
}

const elementoDe = (n) => (n.nodeType === 3 ? n.parentElement : n);

function criarMarca(tag, autor) {
  const el = document.createElement(tag);
  el.className = "rev";
  el.dataset.autor = autor;
  el.dataset.data = agora();
  return el;
}

// Exclusão com controle: texto de outros vira <del>; texto que a própria pessoa
// incluiu (ainda não aceito) é apagado de verdade, como no Word.
export function marcarExclusao(range, autor) {
  let primeiro = null;
  for (const seg of textosNoIntervalo(range)) {
    const ins = seg.parentElement.closest("ins.rev");
    if (ins && ins.dataset.autor === autor) {
      seg.remove();
      if (!ins.textContent) ins.remove();
      continue;
    }
    if (seg.parentElement.closest("del.rev")) { primeiro = primeiro || seg.parentElement.closest("del.rev"); continue; }
    const del = criarMarca("del", autor);
    seg.parentNode.insertBefore(del, seg);
    del.appendChild(seg);
    const junto = juntarVizinhos(del);
    primeiro = primeiro || junto;
  }
  return primeiro;
}

function colocarCursor(no, offset) {
  const r = document.createRange();
  r.setStart(no, offset);
  r.collapse(true);
  const s = document.getSelection();
  s.removeAllRanges();
  s.addRange(r);
}

export function inserirComControle(texto, autor) {
  const s = document.getSelection();
  let r = s.getRangeAt(0);
  // Não se digita dentro de texto excluído: o cursor pula para depois dele.
  const delAtual = elementoDe(r.startContainer).closest("del.rev");
  if (delAtual) { r = document.createRange(); r.setStartAfter(delAtual); r.collapse(true); }
  const insAtual = elementoDe(r.startContainer).closest("ins.rev");
  if (insAtual && insAtual.dataset.autor === autor) {
    if (r.startContainer.nodeType === 3) {
      r.startContainer.insertData(r.startOffset, texto);
      return colocarCursor(r.startContainer, r.startOffset + texto.length);
    }
    const t = document.createTextNode(texto);
    r.insertNode(t);
    return colocarCursor(t, t.length);
  }
  // Logo depois de uma inclusão própria: continua nela.
  const antes = r.startContainer.nodeType === 3
    ? (r.startOffset === 0 ? r.startContainer.previousSibling : null)
    : r.startContainer.childNodes[r.startOffset - 1];
  if (antes?.nodeType === 1 && antes.matches("ins.rev") && antes.dataset.autor === autor) {
    let ult = antes.lastChild;
    if (!ult || ult.nodeType !== 3) { ult = document.createTextNode(""); antes.appendChild(ult); }
    ult.appendData(texto);
    return colocarCursor(ult, ult.length);
  }
  // Texto quebrado em dois se o cursor está no meio dele.
  if (r.startContainer.nodeType === 3 && r.startOffset > 0 && r.startOffset < r.startContainer.length) {
    const resto = r.startContainer.splitText(r.startOffset);
    r = document.createRange();
    r.setStartBefore(resto);
    r.collapse(true);
  }
  const ins = criarMarca("ins", autor);
  const t = document.createTextNode(texto);
  ins.appendChild(t);
  r.insertNode(ins);
  juntarVizinhos(ins);
  colocarCursor(t, t.length);
}

// Trata o evento "beforeinput" do editor quando o controle está ligado.
// Devolve true se a edição foi feita aqui (e o padrão do navegador foi cancelado).
export function tratarEntrada(e, autor) {
  const s = document.getSelection();
  if (!s.rangeCount) return false;
  const tipo = e.inputType;
  const textoDe = () => e.data ?? e.dataTransfer?.getData("text/plain") ?? "";
  if (["insertText", "insertReplacementText", "insertFromPaste", "insertFromDrop", "insertFromYank"].includes(tipo)) {
    e.preventDefault();
    const texto = textoDe().replace(/\r?\n/g, " ");
    if (!s.isCollapsed) {
      const range = s.getRangeAt(0);
      const ref = marcarExclusao(range, autor);
      if (ref?.isConnected) { const r = document.createRange(); r.setStartAfter(ref); r.collapse(true); s.removeAllRanges(); s.addRange(r); }
      else s.collapseToEnd();
    }
    if (texto) inserirComControle(texto, autor);
    return true;
  }
  if (tipo.startsWith("delete")) {
    e.preventDefault();
    const paraTras = /Backward/.test(tipo);
    if (s.isCollapsed) {
      const unidade = /Word/.test(tipo) ? "word" : /Line/.test(tipo) ? "lineboundary" : "character";
      s.modify("extend", paraTras ? "backward" : "forward", unidade);
    }
    if (s.isCollapsed) return true;
    const range = s.getRangeAt(0);
    const inicio = { n: range.startContainer, o: range.startOffset };
    const ref = marcarExclusao(range, autor);
    // Cursor: para trás fica antes do trecho excluído; para frente, depois.
    try {
      if (ref?.isConnected) {
        const r = document.createRange();
        if (paraTras) r.setStartBefore(ref); else r.setStartAfter(ref);
        r.collapse(true);
        s.removeAllRanges();
        s.addRange(r);
      } else if (inicio.n.isConnected) colocarCursor(inicio.n, Math.min(inicio.o, inicio.n.length ?? inicio.o));
    } catch { /* cursor fica onde estiver */ }
    return true;
  }
  return false; // Enter, formatação etc.: segue o comportamento normal
}

// ---------- Aceitar / rejeitar ----------

export function aceitar(el) {
  if (el.tagName === "INS") desembrulhar(el);
  else el.remove();
}
export function rejeitar(el) {
  if (el.tagName === "INS") el.remove();
  else desembrulhar(el);
}

// ---------- Cores por autor ----------

const CORES = ["#1f5fa8", "#a8321f", "#2e7d32", "#7b1fa2", "#c26a00", "#00838f", "#ad1457"];
export function corAutor(nome) {
  let h = 0;
  for (const ch of String(nome)) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return CORES[h % CORES.length];
}
export function cssAutores(autores) {
  return autores.map((a) => {
    const sel = `[data-autor="${String(a).replace(/["\\]/g, "\\$&")}"]`;
    return `.ed-texto ins.rev${sel} { color: ${corAutor(a)}; } .ed-texto del.rev${sel} { color: ${corAutor(a)}; }`;
  }).join("\n");
}
