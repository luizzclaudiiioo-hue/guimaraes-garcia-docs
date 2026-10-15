// Operações de tabela do editor de Texto, respeitando células mescladas
// (colspan/rowspan) como o Word.

const BORDA = "1px solid #000";

// Mapa da tabela: g[linha][coluna] = célula que ocupa aquela posição.
export function grade(tabela) {
  const g = [];
  [...tabela.rows].forEach((tr, r) => {
    g[r] = g[r] || [];
    let c = 0;
    for (const td of tr.cells) {
      while (g[r][c]) c++;
      for (let i = 0; i < td.rowSpan; i++) {
        for (let j = 0; j < td.colSpan; j++) (g[r + i] = g[r + i] || [])[c + j] = td;
      }
      c += td.colSpan;
    }
  });
  return g;
}
export const numColunas = (g) => Math.max(0, ...g.map((l) => l.length));

export function posicao(g, td) {
  for (let r = 0; r < g.length; r++) for (let c = 0; c < (g[r] || []).length; c++) if (g[r][c] === td) return { r, c };
  return null;
}

function novaCelula(modelo) {
  const td = document.createElement("td");
  td.style.cssText = modelo?.style.cssText || `border:${BORDA};padding:2px 5px`;
  td.style.backgroundColor = "";
  td.style.width = "";
  td.innerHTML = "<br>";
  return td;
}

// Coloca a célula na linha certa, antes da primeira célula que começa depois da coluna c.
function inserirNaLinha(tabela, g, r, c, td) {
  const tr = tabela.rows[r];
  const ref = [...tr.cells].find((x) => posicao(g, x).c > c);
  tr.insertBefore(td, ref || null);
}

export function inserirLinha(tabela, td, abaixo) {
  const g = grade(tabela);
  const p = posicao(g, td);
  const alvo = abaixo ? p.r + td.rowSpan : p.r; // a nova linha entra antes desta
  const modelo = g[abaixo ? alvo - 1 : alvo];
  const tr = document.createElement("tr");
  const estendidas = new Set();
  for (let c = 0; c < numColunas(g); c++) {
    const cima = g[alvo - 1]?.[c], baixo = g[alvo]?.[c];
    if (cima && cima === baixo) { // célula mesclada atravessa o ponto de inserção
      if (!estendidas.has(cima)) { cima.rowSpan += 1; estendidas.add(cima); }
      continue;
    }
    tr.appendChild(novaCelula(modelo?.[c]));
  }
  const ref = tabela.rows[alvo];
  if (ref) ref.parentNode.insertBefore(tr, ref);
  else tabela.rows[tabela.rows.length - 1].parentNode.appendChild(tr);
  return tr.cells[0] || td;
}

export function inserirColuna(tabela, td, direita) {
  const g = grade(tabela);
  const p = posicao(g, td);
  const alvo = direita ? p.c + td.colSpan : p.c; // a nova coluna entra antes desta
  const estendidas = new Set();
  const novas = [];
  g.forEach((linha, r) => {
    const esq = linha[alvo - 1], dir = linha[alvo];
    if (esq && esq === dir) {
      if (!estendidas.has(esq)) { esq.colSpan += 1; estendidas.add(esq); }
      return;
    }
    const nova = novaCelula(linha[direita ? alvo - 1 : alvo] || linha[alvo - 1]);
    novas.push([r, nova]);
  });
  for (const [r, nova] of novas) inserirNaLinha(tabela, g, r, alvo - 0.5, nova);
  ajustarColgroup(tabela, alvo, +1);
}

export function excluirLinhas(tabela, td) {
  const p = posicao(grade(tabela), td);
  for (let k = p.r + td.rowSpan - 1; k >= p.r; k--) {
    const g = grade(tabela);
    const vistas = new Set();
    for (let c = 0; c < numColunas(g); c++) {
      const cel = g[k]?.[c];
      if (!cel || vistas.has(cel)) continue;
      vistas.add(cel);
      const inicio = posicao(g, cel);
      if (inicio.r < k) cel.rowSpan -= 1; // vem de cima: só encolhe
      else if (cel.rowSpan > 1 && tabela.rows[k + 1]) { // começa aqui e desce: passa para a linha de baixo
        cel.rowSpan -= 1;
        inserirNaLinha(tabela, g, k + 1, inicio.c - 0.5, cel);
      }
    }
    tabela.rows[k].remove();
  }
  if (!tabela.rows.length) tabela.remove();
}

export function excluirColunas(tabela, td) {
  const p = posicao(grade(tabela), td);
  for (let k = p.c + td.colSpan - 1; k >= p.c; k--) {
    const g = grade(tabela);
    const vistas = new Set();
    for (const linha of g) {
      const cel = linha?.[k];
      if (!cel || vistas.has(cel)) continue;
      vistas.add(cel);
      if (cel.colSpan > 1) cel.colSpan -= 1;
      else cel.remove();
    }
    ajustarColgroup(tabela, k, -1);
  }
  if (!tabela.querySelector("td,th")) tabela.remove();
}

// Retângulo que cobre as duas células por completo (crescendo até englobar as mescladas).
export function retangulo(tabela, a, b) {
  const g = grade(tabela);
  const pa = posicao(g, a), pb = posicao(g, b);
  let r1 = Math.min(pa.r, pb.r), c1 = Math.min(pa.c, pb.c);
  let r2 = Math.max(pa.r + a.rowSpan - 1, pb.r + b.rowSpan - 1), c2 = Math.max(pa.c + a.colSpan - 1, pb.c + b.colSpan - 1);
  for (let mudou = true; mudou;) {
    mudou = false;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const cel = g[r]?.[c];
      if (!cel) continue;
      const p = posicao(g, cel);
      const e = { r1: Math.min(r1, p.r), c1: Math.min(c1, p.c), r2: Math.max(r2, p.r + cel.rowSpan - 1), c2: Math.max(c2, p.c + cel.colSpan - 1) };
      if (e.r1 !== r1 || e.c1 !== c1 || e.r2 !== r2 || e.c2 !== c2) { ({ r1, c1, r2, c2 } = e); mudou = true; }
    }
  }
  const celulas = new Set();
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (g[r]?.[c]) celulas.add(g[r][c]);
  return { r1, c1, r2, c2, celulas: [...celulas] };
}

export function mesclar(tabela, ret) {
  const g = grade(tabela);
  const mestre = g[ret.r1][ret.c1];
  for (const cel of ret.celulas) {
    if (cel === mestre) continue;
    const vazio = !cel.textContent.trim() && !cel.querySelector("img");
    if (!vazio) {
      if (mestre.textContent.trim() || mestre.querySelector("img")) mestre.appendChild(document.createElement("br"));
      else mestre.innerHTML = "";
      mestre.append(...cel.childNodes);
    }
    cel.remove();
  }
  mestre.rowSpan = ret.r2 - ret.r1 + 1;
  mestre.colSpan = ret.c2 - ret.c1 + 1;
  return mestre;
}

export function dividir(tabela, td) {
  if (td.rowSpan === 1 && td.colSpan === 1) return;
  const g = grade(tabela);
  const p = posicao(g, td);
  const { rowSpan: rs, colSpan: cs } = td;
  td.rowSpan = 1;
  td.colSpan = 1;
  for (let r = p.r; r < p.r + rs; r++) {
    const novas = [];
    for (let c = p.c; c < p.c + cs; c++) if (r !== p.r || c !== p.c) novas.push(novaCelula(td));
    if (r === p.r) { td.after(...novas); continue; }
    const tr = tabela.rows[r];
    const ref = [...tr.cells].find((x) => posicao(g, x)?.c > p.c) || null;
    novas.forEach((n) => tr.insertBefore(n, ref));
  }
}

export function bordas(tabela, celulas, tipo) {
  const g = grade(tabela);
  const todas = celulas?.length > 1 ? celulas : [...tabela.querySelectorAll("td,th")];
  const pos = todas.map((td) => ({ td, ...posicao(g, td) }));
  const r1 = Math.min(...pos.map((x) => x.r)), c1 = Math.min(...pos.map((x) => x.c));
  const r2 = Math.max(...pos.map((x) => x.r + x.td.rowSpan - 1)), c2 = Math.max(...pos.map((x) => x.c + x.td.colSpan - 1));
  for (const { td, r, c } of pos) {
    if (tipo === "todas") td.style.border = BORDA;
    else if (tipo === "nenhuma") td.style.border = "none";
    else if (tipo === "externas") {
      td.style.border = "none";
      if (r === r1) td.style.borderTop = BORDA;
      if (r + td.rowSpan - 1 === r2) td.style.borderBottom = BORDA;
      if (c === c1) td.style.borderLeft = BORDA;
      if (c + td.colSpan - 1 === c2) td.style.borderRight = BORDA;
    }
  }
}

// ---------- Largura das colunas (colgroup em %) ----------

export function garantirColgroup(tabela) {
  let cg = tabela.querySelector(":scope > colgroup");
  const g = grade(tabela);
  const n = numColunas(g);
  if (cg && cg.children.length === n) {
    // Larguras vindas do Word (em pontos) viram porcentagem, para arrastar com precisão.
    const cols = [...cg.children];
    if (cols.some((c) => !/%$/.test(c.style.width))) {
      const v = cols.map((c) => parseFloat(c.style.width) || 0);
      const t = v.reduce((a, b) => a + b, 0) || cols.length;
      cols.forEach((c, i) => { c.style.width = `${(((v[i] || t / cols.length) / t) * 100).toFixed(2)}%`; });
    }
    return cg;
  }
  cg?.remove();
  // Mede cada coluna por uma célula que ocupe só ela.
  const larg = Array.from({ length: n }, (_, c) => {
    const cel = g.map((l) => l[c]).find((x) => x && x.colSpan === 1);
    return cel ? cel.getBoundingClientRect().width : 0;
  });
  const total = larg.reduce((a, b) => a + b, 0) || n;
  cg = document.createElement("colgroup");
  for (const w of larg) {
    const col = document.createElement("col");
    col.style.width = `${(((w || total / n) / total) * 100).toFixed(2)}%`;
    cg.appendChild(col);
  }
  tabela.insertBefore(cg, tabela.firstChild);
  tabela.style.tableLayout = "fixed";
  return cg;
}

function ajustarColgroup(tabela, c, delta) {
  const cg = tabela.querySelector(":scope > colgroup");
  if (!cg) return;
  if (delta < 0) cg.children[c]?.remove();
  else {
    const col = document.createElement("col");
    const n = cg.children.length + 1;
    col.style.width = `${(100 / n).toFixed(2)}%`;
    // redistribui proporcionalmente
    [...cg.children].forEach((x) => { x.style.width = `${(parseFloat(x.style.width) * (n - 1) / n).toFixed(2)}%`; });
    cg.insertBefore(col, cg.children[c] || null);
  }
}

// Arrasta a divisa direita da coluna `c`; ajusta só ela e a vizinha.
export function iniciarRedimensionar(tabela, c, xInicial, aoTerminar) {
  const cg = garantirColgroup(tabela);
  const cols = [...cg.children];
  if (c >= cols.length - 1) return false; // borda externa: não redimensiona
  const total = tabela.getBoundingClientRect().width;
  const pct = cols.map((x) => parseFloat(x.style.width));
  const minimo = (20 / total) * 100;
  const mover = (e) => {
    let d = ((e.clientX - xInicial) / total) * 100;
    d = Math.max(minimo - pct[c], Math.min(pct[c + 1] - minimo, d));
    cols[c].style.width = `${(pct[c] + d).toFixed(2)}%`;
    cols[c + 1].style.width = `${(pct[c + 1] - d).toFixed(2)}%`;
  };
  const fim = () => {
    window.removeEventListener("mousemove", mover);
    window.removeEventListener("mouseup", fim);
    document.body.style.cursor = "";
    aoTerminar?.();
  };
  window.addEventListener("mousemove", mover);
  window.addEventListener("mouseup", fim);
  document.body.style.cursor = "col-resize";
  return true;
}

export function novaTabela(linhas, colunas) {
  const td = `<td style="border:${BORDA};padding:2px 5px"><br></td>`;
  const cols = `<colgroup>${`<col style="width:${(100 / colunas).toFixed(2)}%">`.repeat(colunas)}</colgroup>`;
  return `<table style="border-collapse:collapse;width:100%;table-layout:fixed">${cols}<tbody>${`<tr>${td.repeat(colunas)}</tr>`.repeat(linhas)}</tbody></table><p><br></p>`;
}
