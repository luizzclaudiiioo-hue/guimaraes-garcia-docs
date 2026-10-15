import { grade, numColunas, posicao } from "./textoTabelas.js";
import { estilosDe } from "./textoEstilos.js";
import { envolverIntervalo } from "./textoRevisao.js";

// Conversão entre o HTML do editor de Texto e arquivos .docx do Word.

export const CM = 37.7953; // pixels por centímetro na tela
export const PAPEIS = { A4: [21, 29.7], Oficio: [21.6, 33], Carta: [21.59, 27.94] };
export const PAGINA_PADRAO = { tam: "A4", orient: "retrato", margens: { sup: 3, inf: 2, esq: 3, dir: 2 } };
export const FONTE_PADRAO = "Arial Narrow";
export const TAMANHO_PADRAO = 11; // pt

export function dimensoes(pagina = PAGINA_PADRAO) {
  const [w, h] = PAPEIS[pagina.tam] || PAPEIS.A4;
  return pagina.orient === "paisagem" ? [h, w] : [w, h];
}

// ---------- HTML do editor → .docx ----------

const TWIP_CM = 566.93;
const ALINHAMENTO = { left: "LEFT", start: "LEFT", center: "CENTER", right: "RIGHT", end: "RIGHT", justify: "JUSTIFIED" };
const TITULOS = { H1: "HEADING_1", H2: "HEADING_2", H3: "HEADING_3" };
const FONT_TAMANHO = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 }; // execCommand("fontSize")

// Formatos de numeração (como "Lista de vários níveis" do Word): [formato, texto] por nível.
const FORMATOS_LISTA = {
  decimal: [["DECIMAL", "%1."], ["DECIMAL", "%2."], ["DECIMAL", "%3."]],
  juridico: [["DECIMAL", "%1."], ["DECIMAL", "%1.%2."], ["DECIMAL", "%1.%2.%3."], ["DECIMAL", "%1.%2.%3.%4."]],
  alfa: [["DECIMAL", "%1."], ["LOWER_LETTER", "%2)"], ["LOWER_ROMAN", "%3)"]],
  romano: [["UPPER_ROMAN", "%1."], ["UPPER_LETTER", "%2."], ["DECIMAL", "%3."]],
  letras: [["LOWER_LETTER", "%1)"], ["LOWER_ROMAN", "%2)"], ["DECIMAL", "%3."]],
};
const MARCADORES = ["\u2022", "\u25E6", "\u25AA"];

function configNumeracao(d) {
  const recuo = (n, largo) => ({ paragraph: { indent: { left: 720 + 540 * n, hanging: largo ? 360 + 180 * n : 360 }, spacing: { before: 0 } } });
  const config = Object.entries(FORMATOS_LISTA).map(([nome, niveis]) => ({
    reference: `num-${nome}`,
    levels: Array.from({ length: 9 }, (_, n) => {
      const [formato, texto] = niveis[n] || [niveis[niveis.length - 1][0], `%${n + 1}.`];
      const t = nome === "juridico" && n >= niveis.length ? Array.from({ length: n + 1 }, (_, i) => `%${i + 1}`).join(".") + "." : texto;
      return { level: n, format: d.LevelFormat[formato], text: t, alignment: d.AlignmentType.LEFT, style: recuo(n, nome === "juridico") };
    }),
  }));
  config.push({
    reference: "marcadores",
    levels: Array.from({ length: 9 }, (_, n) => ({ level: n, format: d.LevelFormat.BULLET, text: MARCADORES[n % 3], alignment: d.AlignmentType.LEFT, style: recuo(n) })),
  });
  return config;
}

function corHex(css) {
  if (!css) return undefined;
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/.exec(css);
  if (m) return m[4] === "0" ? undefined : [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("");
  const h = /^#([0-9a-f]{6})$/i.exec(css.trim());
  return h ? h[1] : undefined;
}

// Medida CSS → twips (1/20 de ponto).
function twips(css) {
  const m = /^(-?[\d.]+)(pt|cm|mm|px|in|em)?$/.exec(String(css || "").trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  return Math.round(n * ({ pt: 20, cm: TWIP_CM, mm: TWIP_CM / 10, px: 15, in: 1440, em: TAMANHO_PADRAO * 20 }[m[2] || "px"]));
}

function bytesDeDataUrl(src) {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/i.exec(src || "");
  if (!m) return null;
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { tipo: m[1].toLowerCase().replace("jpeg", "jpg"), bytes };
}

// Devolve o .docx como Blob (quem chama decide se grava no arquivo ou baixa).
export async function htmlParaDocx(doc, { html, cabecalho, rodape, cabecalho1 = "", rodape1 = "", notas = {}, comentarios = {} }) {
  const d = await import("docx");
  const { estimatePageNumbersWith } = await import("docx/layout");
  const parse = (h) => new DOMParser().parseFromString(`<div>${h || ""}</div>`, "text/html").body.firstChild;
  const pagina = doc.pagina || PAGINA_PADRAO;
  const est = estilosDe(doc);
  const alinh = (a) => d.AlignmentType[ALINHAMENTO[a] || "LEFT"];
  const [largCm] = dimensoes(pagina);
  const largUtilPx = (largCm - pagina.margens.esq - pagina.margens.dir) * CM;

  function estiloRun(el, base) {
    const s = { ...base };
    const tag = el.tagName;
    if (tag === "B" || tag === "STRONG") s.bold = true;
    if (tag === "I" || tag === "EM") s.italics = true;
    if (tag === "U") s.underline = {};
    if (tag === "S" || tag === "STRIKE") s.strike = true;
    if (tag === "SUP") s.superScript = true;
    if (tag === "SUB") s.subScript = true;
    if (tag === "FONT") {
      if (el.getAttribute("size")) s.size = (FONT_TAMANHO[el.getAttribute("size")] || 12) * 2;
      if (el.getAttribute("color")) s.color = corHex(el.getAttribute("color"));
      if (el.getAttribute("face")) s.font = el.getAttribute("face");
    }
    const st = el.style;
    if (st.fontWeight === "bold" || Number(st.fontWeight) >= 600) s.bold = true;
    if (st.fontWeight === "normal" || (Number(st.fontWeight) && Number(st.fontWeight) < 600)) s.bold = false;
    if (st.fontStyle === "italic") s.italics = true;
    if (st.textDecoration.includes("underline") || st.textDecorationLine.includes("underline")) s.underline = {};
    if (st.textDecoration.includes("line-through") || st.textDecorationLine.includes("line-through")) s.strike = true;
    if (corHex(st.color)) s.color = corHex(st.color);
    if (!/^(P|DIV|H[1-6]|LI|BLOCKQUOTE|TD|TH)$/.test(tag) && corHex(st.backgroundColor)) s.shading ={ type: d.ShadingType.CLEAR, fill: corHex(st.backgroundColor) };
    if (st.fontFamily) s.font = st.fontFamily.split(",")[0].replace(/["']/g, "").trim();
    const t = twips(st.fontSize);
    if (t) s.size = Math.round(t / 10);
    return s;
  }

  function imagem(img) {
    const dados = bytesDeDataUrl(img.getAttribute("src"));
    if (!dados) return null;
    let w = parseFloat(img.getAttribute("width") || img.style.width) || 300;
    let h = parseFloat(img.getAttribute("height") || img.style.height) || w * 0.75;
    if (w > largUtilPx) { h = (h * largUtilPx) / w; w = largUtilPx; }
    const flutua = img.style.float === "left" || img.style.float === "right";
    return new d.ImageRun({
      type: dados.tipo, data: dados.bytes, transformation: { width: Math.round(w), height: Math.round(h) },
      // Imagem com texto ao redor: flutuante, presa ao parágrafo, alinhada à esquerda/direita da coluna.
      floating: flutua ? {
        horizontalPosition: { relative: d.HorizontalPositionRelativeFrom.COLUMN, align: img.style.float === "left" ? d.HorizontalPositionAlign.LEFT : d.HorizontalPositionAlign.RIGHT },
        verticalPosition: { relative: d.VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
        wrap: { type: d.TextWrappingType.SQUARE, side: d.TextWrappingSide.BOTH_SIDES },
        margins: { left: 114300, right: 114300, top: 0, bottom: 57150 },
      } : undefined,
    });
  }

  const BLOCO_RE = /^(P|DIV|H[1-6]|LI|TD|TH|BLOCKQUOTE)$/;
  function temConteudoDepois(no) {
    for (let n = no; n && !(n.nodeType === 1 && BLOCO_RE.test(n.tagName)); n = n.parentNode) {
      for (let s = n.nextSibling; s; s = s.nextSibling) {
        if (s.nodeType === 3 ? s.textContent.length : s.nodeType === 1 && (s.textContent.length || s.querySelector?.("img,br") || /^(IMG|BR)$/.test(s.tagName))) return true;
      }
    }
    return false;
  }

  function runs(no, estilo) {
    const out = [];
    for (const filho of no.childNodes) {
      if (filho.nodeType === 3) {
        // Tabulação e quebras de linha do texto viram os equivalentes do Word.
        const { __rev, ...est } = estilo;
        if (__rev) {
          // Texto incluído/excluído com controle de alterações (autor e data)
          const Classe = __rev.tipo === "ins" ? d.InsertedTextRun : d.DeletedTextRun;
          const texto = filho.textContent.replace(/[\t\n]/g, " ");
          if (texto) out.push(new Classe({ text: texto, ...est, id: revId++, author: __rev.autor, date: __rev.data }));
          continue;
        }
        const partes = filho.textContent.replace(/\u200b/g, "").split("\t");
        partes.forEach((p, i) => {
          if (i > 0) out.push(new d.TextRun({ children: [new d.Tab()], ...est }));
          if (p) out.push(new d.TextRun({ text: p.replace(/\n/g, " "), ...est }));
        });
        continue;
      }
      if (filho.nodeType !== 1) continue;
      const tag = filho.tagName;
      if (tag === "BR") {
        // O <br> no fim do parágrafo só marca linha vazia no HTML; no Word viraria uma linha a mais.
        if (!temConteudoDepois(filho)) continue;
        const { __rev, ...est } = estilo;
        out.push(new d.TextRun({ break: 1, ...est }));
        continue;
      }
      if (tag === "UL" || tag === "OL") continue; // sublista: vira parágrafos próprios
      if (tag === "IMG") { const im = imagem(filho); if (im) out.push(im); continue; }
      if (tag === "SUP" && filho.classList.contains("nota")) {
        const n = numeroNota.get(filho.dataset.nota);
        if (n) out.push(new d.FootnoteReferenceRun(n));
        continue;
      }
      if (filho.dataset?.campo === "pagina") { out.push(new d.TextRun({ children: [d.PageNumber.CURRENT], ...estilo })); continue; }
      if (filho.dataset?.campo === "paginas") { out.push(new d.TextRun({ children: [d.PageNumber.TOTAL_PAGES], ...estilo })); continue; }
      if (tag === "A" && filho.getAttribute("href")) {
        out.push(new d.ExternalHyperlink({ link: filho.getAttribute("href"), children: runs(filho, { ...estiloRun(filho, estilo), style: "Hyperlink" }) }));
        continue;
      }
      if ((tag === "INS" || tag === "DEL") && filho.classList.contains("rev")) {
        out.push(...runs(filho, { ...estiloRun(filho, estilo), __rev: { tipo: tag.toLowerCase(), autor: filho.dataset.autor || "Autor", data: filho.dataset.data || new Date().toISOString() } }));
        continue;
      }
      if (tag === "SPAN" && filho.classList.contains("coment") && numeroCom.has(filho.dataset.c)) {
        // Trecho comentado: marca início/fim no primeiro e no último pedaço com o mesmo id.
        const n = numeroCom.get(filho.dataset.c);
        const a = ancoras.get(filho.dataset.c);
        if (a.primeiro === filho) out.push(new d.CommentRangeStart(n));
        out.push(...runs(filho, estiloRun(filho, estilo)));
        if (a.ultimo === filho) out.push(new d.CommentRangeEnd(n), new d.TextRun({ children: [new d.CommentReference(n)] }));
        continue;
      }
      out.push(...runs(filho, estiloRun(filho, estilo)));
    }
    return out;
  }

  function propsParagrafo(el) {
    const st = el.style || {};
    const p = {};
    const alin = ALINHAMENTO[st.textAlign || el.getAttribute?.("align") || ""];
    if (alin) p.alignment = d.AlignmentType[alin];
    const sp = {};
    // O tipo vai sempre explícito: sem ele o LibreOffice entende "exatamente" e corta imagens.
    if (st.lineHeight === "normal") { sp.line = 240; sp.lineRule = d.LineRuleType.AUTO; } // "simples" do Word
    else if (st.lineHeight) {
      if (/^[\d.]+$/.test(st.lineHeight)) { sp.line = Math.round(parseFloat(st.lineHeight) * 240); sp.lineRule = d.LineRuleType.AUTO; }
      else { const t = twips(st.lineHeight); if (t) { sp.line = t; sp.lineRule = d.LineRuleType.AT_LEAST; } }
    }
    // Parágrafo com imagem e sem entrelinhas definida: simples (1,5 cortaria a imagem).
    if (!sp.line && el.querySelector?.("img")) { sp.line = 240; sp.lineRule = d.LineRuleType.AUTO; }
    if (st.marginTop) { const t = twips(st.marginTop); if (t !== null) sp.before = Math.max(0, t); }
    if (st.marginBottom) { const t = twips(st.marginBottom); if (t !== null) sp.after = Math.max(0, t); }
    if (Object.keys(sp).length) p.spacing = sp;
    const ind = {};
    if (st.textIndent) { const t = twips(st.textIndent); if (t > 0) ind.firstLine = t; else if (t < 0) ind.hanging = -t; }
    if (st.marginLeft) { const t = twips(st.marginLeft); if (t > 0) ind.left = t; }
    if (st.marginRight) { const t = twips(st.marginRight); if (t > 0) ind.right = t; }
    if (Object.keys(ind).length) p.indent = ind;
    const bordas = {};
    if (st.borderTop && !/none/.test(st.borderTop)) bordas.top = { ...borda(st.borderTop), space: 1 };
    if (st.borderBottom && !/none/.test(st.borderBottom)) bordas.bottom = { ...borda(st.borderBottom), space: 1 };
    if (st.borderLeft && !/none/.test(st.borderLeft)) bordas.left = { ...borda(st.borderLeft), space: 4 };
    if (st.borderRight && !/none/.test(st.borderRight)) bordas.right = { ...borda(st.borderRight), space: 4 };
    if (Object.keys(bordas).length) p.border = bordas;
    const tabs = (el.dataset?.tabs || "").split(",").map(Number).filter((v) => v > 0);
    if (tabs.length) p.tabStops = tabs.map((v) => ({ type: d.TabStopType.LEFT, position: Math.round(v * TWIP_CM) }));
    const fundo = corHex(st.backgroundColor);
    if (fundo) p.shading = { type: d.ShadingType.CLEAR, color: "auto", fill: fundo };
    return p;
  }

  function paragrafo(el, extra = {}) {
    if (el.tagName === "BLOCKQUOTE" || el.closest?.("blockquote")) extra = { style: "Citacao", ...extra };
    const base = estiloRun(el, {});
    // Fonte da marca de parágrafo: define a altura de linhas em branco no Word.
    const marca = {};
    if (base.font) marca.font = base.font;
    if (base.size) marca.size = base.size;
    const props = propsParagrafo(el);
    return new d.Paragraph({
      children: runs(el, base),
      run: Object.keys(marca).length ? marca : undefined,
      heading: TITULOS[el.tagName] ? d.HeadingLevel[TITULOS[el.tagName]] : undefined,
      ...props,
      ...extra,
    });
  }

  function borda(css) {
    const m = /([\d.]+)px\s+(\w+)\s+(.*)/.exec(css || "");
    if (!m || m[2] === "none" || parseFloat(m[1]) === 0) return { style: d.BorderStyle.NONE, size: 0, color: "FFFFFF" };
    return { style: m[2] === "dashed" ? d.BorderStyle.DASHED : m[2] === "dotted" ? d.BorderStyle.DOTTED : d.BorderStyle.SINGLE, size: Math.max(2, Math.round(parseFloat(m[1]) * 6)), color: corHex(m[3]) || "000000" };
  }

  function tabela(el) {
    // Larguras: <colgroup> (em % ou pontos) ou colunas iguais; tudo em twips na largura útil.
    const g = grade(el);
    const n = numColunas(g) || 1;
    const larguraUtil = Math.round((largCm - pagina.margens.esq - pagina.margens.dir) * TWIP_CM);
    const totalTab = /%$/.test(el.style.width) ? Math.round((larguraUtil * parseFloat(el.style.width)) / 100) : larguraUtil;
    const cols = [...(el.querySelector(":scope > colgroup")?.children || [])].map((c) => parseFloat(c.style.width || c.getAttribute("width")) || 0);
    const soma = cols.reduce((a, b) => a + b, 0);
    const larg = cols.length === n && soma ? cols.map((w) => Math.round((w / soma) * totalTab)) : Array(n).fill(Math.round(totalTab / n));
    const linhas = [...el.rows].map((tr) => new d.TableRow({
      children: [...tr.cells].map((td) => {
        const conteudo = blocos(td);
        const st = td.style;
        const c0 = posicao(g, td)?.c ?? 0;
        const w = larg.slice(c0, c0 + td.colSpan).reduce((a, b) => a + b, 0);
        return new d.TableCell({
          children: conteudo.length ? conteudo : [new d.Paragraph("")],
          columnSpan: td.colSpan > 1 ? td.colSpan : undefined,
          rowSpan: td.rowSpan > 1 ? td.rowSpan : undefined,
          width: w ? { size: w, type: d.WidthType.DXA } : undefined,
          verticalAlign: st.verticalAlign === "middle" ? d.VerticalAlign.CENTER : st.verticalAlign === "bottom" ? d.VerticalAlign.BOTTOM : d.VerticalAlign.TOP,
          shading: corHex(st.backgroundColor) ? { type: d.ShadingType.CLEAR, fill: corHex(st.backgroundColor) } : undefined,
          // Sem borda definida no HTML = sem borda no Word (o padrão do Word seria com borda).
          borders: {
            top: borda(st.borderTop || st.border), bottom: borda(st.borderBottom || st.border),
            left: borda(st.borderLeft || st.border), right: borda(st.borderRight || st.border),
          },
        });
      }),
    }));
    return new d.Table({ rows: linhas, width: { size: totalTab, type: d.WidthType.DXA }, columnWidths: larg, layout: d.TableLayoutType.FIXED });
  }

  // Lista com subníveis (Tab): o Chrome põe a sublista dentro do <li> ou direto no <ol>.
  let instanciaLista = 0; // cada lista recomeça do 1
  function lista(el, nivel, ref, inst, out) {
    const r = el.tagName === "OL" ? ref : "marcadores";
    for (const filho of el.children) {
      if (filho.tagName === "LI") {
        out.push(paragrafo(filho, { numbering: { reference: r, level: Math.min(nivel, 8), instance: inst } }));
        for (const sub of filho.children) if (sub.tagName === "OL" || sub.tagName === "UL") lista(sub, nivel + 1, ref, inst, out);
      } else if (filho.tagName === "OL" || filho.tagName === "UL") lista(filho, nivel + 1, ref, inst, out);
    }
  }

  const BLOCO = ["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE"];
  function blocos(no) {
    const out = [];
    let soltos = null; // texto e elementos de linha fora de parágrafo
    const fecharSoltos = () => { if (soltos) { out.push(paragrafo(soltos)); soltos = null; } };
    for (const filho of no.childNodes) {
      const tag = filho.nodeType === 1 ? filho.tagName : null;
      if (tag && BLOCO.includes(tag) && !filho.dataset?.sumario) {
        fecharSoltos();
        if ([...filho.children].some((c) => [...BLOCO, "UL", "OL", "TABLE"].includes(c.tagName))) out.push(...blocos(filho));
        else out.push(paragrafo(filho));
      } else if (tag === "UL" || tag === "OL") {
        fecharSoltos();
        lista(filho, 0, `num-${(tag === "OL" && FORMATOS_LISTA[filho.dataset.num] && filho.dataset.num) || "decimal"}`, ++instanciaLista, out);
      } else if (tag === "TABLE") {
        fecharSoltos();
        out.push(tabela(filho));
      } else if (tag === "DIV" && filho.dataset?.sumario) {
        // Sumário: campo do Word (ele calcula os números de página ao abrir).
        fecharSoltos();
        temSumario = true;
        out.push(new d.Paragraph({ alignment: d.AlignmentType.CENTER, spacing: { after: 120 }, indent: { firstLine: 0 }, children: [new d.TextRun({ text: "Sumário", bold: true, size: 28 })] }));
        out.push(new d.TableOfContents("Sumário", { hyperlink: true, headingStyleRange: "1-3" }));
      } else if (tag === "HR") {
        fecharSoltos();
        out.push(new d.Paragraph({ children: [new d.PageBreak()] }));
      } else if (filho.nodeType === 3 && !filho.textContent.trim() && !soltos) {
        continue;
      } else {
        if (!soltos) soltos = document.createElement("p");
        soltos.appendChild(filho.cloneNode(true));
      }
    }
    fecharSoltos();
    return out;
  }

  const [w, h] = PAPEIS[pagina.tam] || PAPEIS.A4;
  const m = pagina.margens;
  // Notas de rodapé, numeradas na ordem em que aparecem no texto.
  const corpo = parse(html);
  const numeroNota = new Map();
  corpo.querySelectorAll("sup.nota[data-nota]").forEach((s) => { if (!numeroNota.has(s.dataset.nota)) numeroNota.set(s.dataset.nota, numeroNota.size + 1); });
  const notasDocx = {};
  let temSumario = false;
  let revId = 1;
  // Comentários: ids numéricos e o primeiro/último pedaço de texto de cada um.
  const ancoras = new Map();
  corpo.querySelectorAll("span.coment[data-c]").forEach((el) => {
    const a = ancoras.get(el.dataset.c) || { primeiro: el };
    a.ultimo = el;
    ancoras.set(el.dataset.c, a);
  });
  const numeroCom = new Map([...ancoras.keys()].filter((id) => comentarios[id]).map((id, i) => [id, i]));
  const rodapeFilhos = blocos(parse(rodape));
  if (doc.numPag) {
    const pagTxt = doc.numPag === "xdey"
      ? [new d.TextRun({ children: ["Página ", d.PageNumber.CURRENT, " de ", d.PageNumber.TOTAL_PAGES] })]
      : [new d.TextRun({ children: [d.PageNumber.CURRENT] })];
    rodapeFilhos.push(new d.Paragraph({ alignment: doc.numPag === "centro" || doc.numPag === "xdey" ? d.AlignmentType.CENTER : d.AlignmentType.RIGHT, children: pagTxt }));
  }
  const cabFilhos = blocos(parse(cabecalho));
  // Primeira página diferente: o Word usa o "first" na 1ª página e o "default" nas demais.
  const cab1 = doc.primeiraDiferente ? blocos(parse(cabecalho1)) : [];
  const rod1 = doc.primeiraDiferente ? blocos(parse(rodape1)) : [];
  const vazio = () => [new d.Paragraph("")];

  const titulo = (id, nome, t) => ({
    id, name: nome, basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: t.tamanho * 2, bold: !!t.negrito, italics: !!t.italico, color: (t.cor || "#000000").replace("#", ""), font: t.fonte },
    paragraph: { spacing: { before: t.antes * 20, after: t.depois * 20, line: 288, lineRule: d.LineRuleType.AUTO }, keepNext: true, alignment: alinh(t.alinhamento), indent: { firstLine: 0 } },
  });
  const citacao = {
    id: "Citacao", name: "Citação", basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: est.citacao.tamanho * 2, font: est.citacao.fonte, italics: !!est.citacao.italico },
    paragraph: {
      spacing: { before: est.citacao.antes * 20, after: est.citacao.depois * 20, line: Math.round(est.citacao.entrelinhas * 240), lineRule: d.LineRuleType.AUTO },
      indent: { left: Math.round(est.citacao.recuoEsq * TWIP_CM), firstLine: 0 }, alignment: alinh(est.citacao.alinhamento),
    },
  };
  // Quebras de seção (<hr class="secao" data-orient>): cada trecho vira uma seção do Word,
  // com a orientação própria; cabeçalho e rodapé continuam os da primeira (o Word herda).
  const partes = [{ orient: pagina.orient === "paisagem" ? "paisagem" : "retrato", el: corpo.ownerDocument.createElement("div") }];
  for (const n of [...corpo.childNodes]) {
    if (n.nodeType === 1 && n.matches("hr.secao")) partes.push({ orient: n.dataset.orient === "paisagem" ? "paisagem" : "retrato", el: corpo.ownerDocument.createElement("div") });
    else partes[partes.length - 1].el.appendChild(n);
  }
  const filhosPartes = partes.map((pt) => blocos(pt.el));
  for (const [id, n] of numeroNota) {
    const b = blocos(parse(notas[id] || ""));
    notasDocx[n] = { children: b.length ? b : [new d.Paragraph("")] };
  }
  const documento = new d.Document({
    numbering: { config: configNumeracao(d) },
    styles: {
      default: { document: {
        run: { font: est.normal.fonte, size: est.normal.tamanho * 2 },
        paragraph: { spacing: { line: Math.round(est.normal.entrelinhas * 240), lineRule: d.LineRuleType.AUTO, before: est.normal.antes * 20, after: est.normal.depois * 20 } },
      } },
      paragraphStyles: [
        // Estilo Normal com recuo e alinhamento (no "padrão do documento" o LibreOffice ignora o alinhamento).
        {
          id: "Normal", name: "Normal", quickFormat: true,
          run: { font: est.normal.fonte, size: est.normal.tamanho * 2 },
          paragraph: {
            spacing: { line: Math.round(est.normal.entrelinhas * 240), lineRule: d.LineRuleType.AUTO, before: est.normal.antes * 20, after: est.normal.depois * 20 },
            indent: { firstLine: Math.round((est.normal.recuo || 0) * TWIP_CM) },
            alignment: alinh(est.normal.alinhamento),
          },
        },
        // Linhas do sumário sem o recuo de 1ª linha do Normal.
        ...[1, 2, 3].map((n) => ({ id: `TOC${n}`, name: `toc ${n}`, basedOn: "Normal", next: "Normal", paragraph: { indent: { firstLine: 0, left: (n - 1) * 240 }, alignment: d.AlignmentType.LEFT, spacing: { after: 60 } } })),
        titulo("Heading1", "Heading 1", est.h1), titulo("Heading2", "Heading 2", est.h2), titulo("Heading3", "Heading 3", est.h3), citacao],
    },
    footnotes: numeroNota.size ? notasDocx : undefined,
    comments: numeroCom.size ? {
      children: [...numeroCom].map(([id, n]) => {
        const c = comentarios[id];
        const linhas = [(c.resolvido ? "[Resolvido] " : "") + (c.texto || ""), ...(c.respostas || []).map((r) => `↳ ${r.autor}: ${r.texto}`)];
        return {
          id: n, author: c.autor || "Autor", initials: (c.autor || "A").split(/\s+/).map((x) => x[0]).join("").slice(0, 3).toUpperCase(),
          date: c.data ? new Date(c.data) : new Date(),
          children: linhas.flatMap((l) => l.split("\n")).map((l) => new d.Paragraph(l)),
        };
      }),
    } : undefined,
    // Sumário já gravado com os números de página calculados (o Word não precisa perguntar
    // se deve atualizar os campos ao abrir; clique direito > Atualizar campo recalcula).
    // Páginas do sumário: medidas com as fontes do navegador (a Arial Narrow não está nas
    // tabelas da biblioteca); guess segue estimando onde o cálculo exato não alcança.
    pageNumbers: temSumario ? estimatePageNumbersWith({ measureWidth: await medidorDeTexto(html, est), guess: true }) : undefined,
    sections: partes.map((pt, i) => ({
      properties: {
        titlePage: i === 0 && !!doc.primeiraDiferente,
        page: {
          size: { width: Math.round(w * TWIP_CM), height: Math.round(h * TWIP_CM), orientation: pt.orient === "paisagem" ? d.PageOrientation.LANDSCAPE : d.PageOrientation.PORTRAIT },
          margin: { top: Math.round(m.sup * TWIP_CM), bottom: Math.round(m.inf * TWIP_CM), left: Math.round(m.esq * TWIP_CM), right: Math.round(m.dir * TWIP_CM), header: Math.round(1.25 * TWIP_CM), footer: Math.round(1.25 * TWIP_CM) },
        },
      },
      ...(i > 0 ? { children: filhosPartes[i].length ? filhosPartes[i] : [new d.Paragraph("")] } : {
      headers: doc.primeiraDiferente
        ? { default: new d.Header({ children: cabFilhos.length ? cabFilhos : vazio() }), first: new d.Header({ children: cab1.length ? cab1 : vazio() }) }
        : cabFilhos.length ? { default: new d.Header({ children: cabFilhos }) } : undefined,
      footers: doc.primeiraDiferente
        ? { default: new d.Footer({ children: rodapeFilhos.length ? rodapeFilhos : vazio() }), first: new d.Footer({ children: rod1.length ? rod1 : vazio() }) }
        : rodapeFilhos.length ? { default: new d.Footer({ children: rodapeFilhos }) } : undefined,
      children: filhosPartes[0].length ? filhosPartes[0] : [new d.Paragraph("")],
      }),
    })),
  });
  return d.Packer.toBlob(documento);
}

// ---------- .docx → HTML do editor ----------

// O docx-preview desenha o documento como o Word (fontes, espaçamentos, numeração,
// cabeçalho, rodapé, imagens). Depois copiamos os estilos calculados pelo navegador
// para dentro do HTML, que é o formato que o editor e o exportador entendem.

const pxPt = (px) => Math.round(parseFloat(px) * 0.75 * 2) / 2;
const pxCm = (px) => Math.round((parseFloat(px) / CM) * 100) / 100;
const ROMANOS = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
function formatarContador(n, estilo) {
  switch (estilo) {
    case "lower-alpha": case "lower-latin": return String.fromCharCode(96 + ((n - 1) % 26) + 1);
    case "upper-alpha": case "upper-latin": return String.fromCharCode(64 + ((n - 1) % 26) + 1);
    case "lower-roman": case "upper-roman": {
      let r = "", x = n;
      for (const [v, s] of ROMANOS) while (x >= v) { r += s; x -= v; }
      return estilo === "upper-roman" ? r.toUpperCase() : r;
    }
    case "decimal-leading-zero": return String(n).padStart(2, "0");
    default: return String(n);
  }
}

// Numeração automática do Word vem como contador CSS em ::before; aqui vira texto.
function criarNumerador() {
  const contagem = {};
  return (el) => {
    const cls = [...el.classList].find((c) => /-num-\d+-\d+$/.test(c));
    if (!cls) return "";
    const [, prefixo, id, nivel] = /^(.*-num-)(\d+)-(\d+)$/.exec(cls);
    const lv = Number(nivel);
    contagem[id] = contagem[id] || [];
    contagem[id][lv] = (contagem[id][lv] || 0) + 1;
    contagem[id].length = lv + 1;
    const conteudo = getComputedStyle(el, "::before").content;
    if (!conteudo || conteudo === "none" || conteudo === "normal") return "";
    let texto = "";
    const re = /"((?:[^"\\]|\\.)*)"|counter\(([^,)]+)(?:,\s*([\w-]+))?\)/g;
    let m;
    while ((m = re.exec(conteudo))) {
      if (m[1] !== undefined) texto += m[1].replace(/\\([0-9a-f]{1,6})\s?/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/\\(.)/g, "$1");
      else {
        const nome = m[2].trim();
        const ml = new RegExp("^" + prefixo.replace(/[-]/g, "\\-") + id + "-(\\d+)$").exec(nome);
        const n = ml ? contagem[id][Number(ml[1])] || 1 : 1;
        texto += formatarContador(n, m[3]);
      }
    }
    // Marcadores em fontes de símbolo (Symbol/Wingdings) viram "•".
    if (/^[-]/.test(texto) || (texto.trim().length === 1 && !/[\w\d]/.test(texto.trim()))) texto = "•";
    return texto;
  };
}

function inlinarEstilos(orig, copia) {
  const os = [orig, ...orig.querySelectorAll("*")];
  const cs = [copia, ...copia.querySelectorAll("*")];
  const numerar = criarNumerador();
  let ultimaFonte = null;
  os.forEach((o, i) => {
    const c = cs[i];
    if (!c) return;
    if (o.dataset?.notaId) { // chamada de nota de rodapé
      const sup = document.createElement("sup");
      sup.className = "nota";
      sup.dataset.nota = o.dataset.notaId;
      sup.setAttribute("contenteditable", "false");
      sup.textContent = "•";
      c.replaceWith(sup);
      return;
    }
    const st = getComputedStyle(o);
    // Continuação de célula mesclada na vertical: o docx-preview só esconde; aqui sai.
    if ((o.tagName === "TD" || o.tagName === "TH") && st.display === "none") { c.remove(); return; }
    const pai = o.parentElement ? getComputedStyle(o.parentElement) : null;
    const decl = [];
    const tag = o.tagName;
    const bloco = /^(P|H[1-6]|LI|TD|TH)$/.test(tag);
    const fsPx = parseFloat(st.fontSize);
    if (bloco) {
      if (["center", "right", "justify"].includes(st.textAlign)) decl.push(`text-align:${st.textAlign}`);
      // Entrelinhas: "múltiplo" vira proporção; "pelo menos"/"exatamente" vira pontos.
      const lhInline = o.style.lineHeight || "";
      const peloMenos = /calc\(100% \+ ([\d.]+)px\)/.exec(lhInline);
      if (peloMenos) decl.push(`line-height:${pxPt(Math.max(parseFloat(peloMenos[1]), fsPx * 1.15))}pt`);
      else if (/^[\d.]+(px|pt)$/.test(lhInline)) decl.push(`line-height:${pxPt(parseFloat(st.lineHeight))}pt`);
      else if (st.lineHeight === "normal") decl.push("line-height:normal");
      else if (fsPx) {
        const r = Math.round((parseFloat(st.lineHeight) / fsPx) * 100) / 100;
        if (r) decl.push(`line-height:${r}`);
      }
      if (tag !== "TD" && tag !== "TH") {
        decl.push(`margin-top:${pxPt(st.marginTop)}pt`, `margin-bottom:${pxPt(st.marginBottom)}pt`);
        if (parseFloat(st.textIndent)) decl.push(`text-indent:${pxCm(st.textIndent)}cm`);
        for (const lado of ["Top", "Bottom", "Left", "Right"]) {
          const estilo = st[`border${lado}Style`];
          if (estilo && estilo !== "none" && parseFloat(st[`border${lado}Width`]) > 0) decl.push(`border-${lado.toLowerCase()}:${st[`border${lado}Width`]} ${estilo} ${st[`border${lado}Color`]}`, `padding-${lado.toLowerCase()}:${lado === "Top" || lado === "Bottom" ? "2pt" : "5pt"}`);
        }
        if (bloco && tag !== "LI" && corHex(st.backgroundColor)) decl.push(`background-color:${st.backgroundColor}`);
        if (parseFloat(st.marginLeft) > 0) decl.push(`margin-left:${pxCm(st.marginLeft)}cm`);
        else if (parseFloat(st.paddingLeft) > 0) decl.push(`margin-left:${pxCm(st.paddingLeft)}cm`);
        if (parseFloat(st.marginRight) > 0) decl.push(`margin-right:${pxCm(st.marginRight)}cm`);
      } else {
        for (const lado of ["Top", "Bottom", "Left", "Right"]) {
          const estilo = st[`border${lado}Style`];
          if (estilo && estilo !== "none" && parseFloat(st[`border${lado}Width`]) > 0) decl.push(`border-${lado.toLowerCase()}:${st[`border${lado}Width`]} ${estilo} ${st[`border${lado}Color`]}`);
        }
        decl.push("padding:2px 5px", "vertical-align:top");
        if (corHex(st.backgroundColor)) decl.push(`background-color:${st.backgroundColor}`);
      }
    }
    // Fontes: no bloco compara com o padrão do editor; dentro do texto, com o elemento pai.
    const familia = st.fontFamily.split(",")[0].replace(/["']/g, "").trim();
    const familiaPai = bloco || !pai ? FONTE_PADRAO : pai.fontFamily.split(",")[0].replace(/["']/g, "").trim();
    if (familia && familia !== familiaPai) decl.push(`font-family:'${familia}'`);
    const pt = pxPt(st.fontSize);
    const ptPai = bloco || !pai ? TAMANHO_PADRAO : pxPt(pai.fontSize);
    if (pt && pt !== ptPai) decl.push(`font-size:${pt}pt`);
    const negrito = Number(st.fontWeight) >= 600, negritoPai = !bloco && pai ? Number(pai.fontWeight) >= 600 : false;
    if (negrito !== negritoPai && !/^H[1-6]$/.test(tag)) decl.push(`font-weight:${negrito ? "bold" : "normal"}`);
    if (st.fontStyle === "italic" && (bloco || pai?.fontStyle !== "italic")) decl.push("font-style:italic");
    const deco = st.textDecorationLine || "";
    const decoPai = !bloco && pai ? pai.textDecorationLine || "" : "";
    const linhas = ["underline", "line-through"].filter((x) => deco.includes(x) && !decoPai.includes(x));
    if (linhas.length) decl.push(`text-decoration:${linhas.join(" ")}`);
    const cor = corHex(st.color);
    if (cor && cor !== (bloco || !pai ? "000000" : corHex(pai.color))) decl.push(`color:#${cor}`);
    if (!bloco && corHex(st.backgroundColor)) decl.push(`background-color:#${corHex(st.backgroundColor)}`);
    const nivelTitulo = tag === "P" && /(?:^|\s)docx_(?:heading|t[ií]?tulo)([1-3])(?:\s|$)/i.exec(o.className || "");
    for (const a of [...c.attributes]) if (!["href", "src", "colspan", "rowspan", "data-com-ini", "data-com-fim", "data-sumario"].includes(a.name)) c.removeAttribute(a.name);
    if (tag === "INS" || tag === "DEL") {
      // Riscado/sublinhado e cor das marcações vêm do editor, não do arquivo.
      for (let k = decl.length - 1; k >= 0; k--) if (/^(text-decoration|color):/.test(decl[k])) decl.splice(k, 1);
      c.className = "rev";
      c.dataset.autor = o.dataset.autor || "Autor desconhecido";
      c.dataset.data = o.dataset.data || o.getAttribute("datetime") || "";
    }
    if (tag === "IMG") {
      // Imagem com texto ao redor: o docx-preview põe o "float" no bloco em volta.
      const flutua = o.parentElement && getComputedStyle(o.parentElement).float;
      if (flutua === "left" || flutua === "right") decl.push(`float:${flutua}`, flutua === "left" ? "margin:0 12px 6px 0" : "margin:0 0 6px 12px");
      c.setAttribute("width", Math.round(o.getBoundingClientRect().width) || o.naturalWidth || 300);
      c.setAttribute("height", Math.round(o.getBoundingClientRect().height) || o.naturalHeight || 200);
    }
    if (tag === "TABLE") decl.push("border-collapse:collapse", "width:100%", "table-layout:fixed");
    if (tag === "COL" && (o.style.width || o.getAttribute("width"))) decl.push(`width:${o.style.width || o.getAttribute("width") + "px"}`);
    if (decl.length) c.setAttribute("style", decl.join(";"));
    // Parágrafo com estilo "Título 1/2/3" do Word vira título de verdade (o sumário depende disso).
    if (nivelTitulo) {
      const h = document.createElement("h" + nivelTitulo[1]);
      if (c.getAttribute("style")) h.setAttribute("style", c.getAttribute("style"));
      h.append(...c.childNodes);
      c.replaceWith(h);
    }
    if (bloco && tag !== "TD" && tag !== "TH") {
      // No Word a altura de uma linha em branco vem da fonte do parágrafo; o HTML não traz
      // essa informação, então usamos a do último texto visto.
      const run = o.querySelector("span") || o;
      if (o.textContent.trim()) ultimaFonte = { fam: getComputedStyle(run).fontFamily.split(",")[0].replace(/["']/g, "").trim(), pt: pxPt(getComputedStyle(run).fontSize) };
      else if (ultimaFonte && !o.querySelector("img")) {
        if (!decl.some((x) => x.startsWith("font-family")) && ultimaFonte.fam !== FONTE_PADRAO) decl.push(`font-family:'${ultimaFonte.fam}'`);
        if (!decl.some((x) => x.startsWith("font-size")) && ultimaFonte.pt !== TAMANHO_PADRAO) decl.push(`font-size:${ultimaFonte.pt}pt`);
        c.setAttribute("style", decl.join(";"));
      }
    }
    if (bloco) {
      const numero = numerar(o);
      if (numero) { c.insertBefore(spanTab(), c.firstChild); c.insertBefore(document.createTextNode(numero), c.firstChild); }
    }
    if (o.classList.contains?.("docx-tab-stop") || /tab-stop/.test(o.className)) {
      // Tabulação: guarda onde o texto seguinte começa como parada do parágrafo (data-tabs,
      // em cm a partir da margem ou da borda da célula), para ficar alinhado como no Word.
      const cel = o.closest("td,th");
      const refEl = cel || o.closest("article") || o.closest("section");
      const refX = refEl ? refEl.getBoundingClientRect().left + (cel ? parseFloat(getComputedStyle(cel).paddingLeft) || 0 : 0) : 0;
      let parada = Math.round(((o.getBoundingClientRect().right - refX) / CM) * 100) / 100;
      if (Math.abs(parada - Math.round(parada * 4) / 4) < 0.1) parada = Math.round(parada * 4) / 4; // o docx-preview arredonda em pontos
      const bloco = c.parentElement?.closest("p,h1,h2,h3,h4,h5,h6,li,td,th");
      if (bloco && parada > 0.1 && !bloco.matches("td,th")) {
        const lista = (bloco.dataset.tabs || "").split(",").map(Number).filter((v) => v > 0 && Math.abs(v - parada) > 0.05);
        bloco.dataset.tabs = [...lista, parada].sort((a, b) => a - b).join(",");
      }
      c.replaceWith(spanTab());
    }
  });
}

function spanTab() {
  const t = document.createElement("span");
  t.className = "tab";
  t.setAttribute("contenteditable", "false");
  t.textContent = "\t";
  return t;
}

function limpar(raiz) {
  // Tira embrulhos vazios (span sem estilo) e elementos que o editor não usa.
  raiz.querySelectorAll("script,style,link,meta").forEach((n) => n.remove());
  raiz.querySelectorAll("p div, h1 div, h2 div, h3 div, h4 div, li div").forEach((n) => {
    const span = document.createElement("span");
    span.append(...n.childNodes);
    n.replaceWith(span);
  });
  raiz.querySelectorAll("span:not([style]):not([data-com-ini]):not([data-com-fim]):not(.tab),font:not([style])").forEach((n) => n.replaceWith(...n.childNodes));
  raiz.querySelectorAll("p").forEach((p) => { if (!p.textContent && !p.querySelector("img,br")) p.innerHTML = "<br>"; });
  return raiz.innerHTML;
}

async function numeroDePagina(buffer) {
  try {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(buffer);
    for (const nome of Object.keys(zip.files).filter((n) => /^word\/footer\d*\.xml$/.test(n))) {
      const xml = await zip.file(nome).async("string");
      if (!/(instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b)/.test(xml)) continue;
      if (/NUMPAGES/.test(xml)) return "xdey";
      const p = (/<w:p[ >](?:(?!<\/w:p>)[\s\S])*?PAGE[\s\S]*?<\/w:p>/.exec(xml) || [""])[0];
      return /<w:jc w:val="center"/.test(p) ? "centro" : "direita";
    }
  } catch { /* sem número */ }
  return "";
}

async function cabecalhosSemTitlePg(buffer, renderAsync) {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("string");
  // Só a primeira seção importa (é a que o editor usa).
  const sectPr = /<w:sectPr[\s>][\s\S]*?<\/w:sectPr>/.exec(xml || "");
  if (!sectPr || !/<w:titlePg(?:\s*\/>|\s+w:val="(?:1|true|on)"\s*\/>)/.test(sectPr[0])) return null;
  zip.file("word/document.xml", xml.replace(/<w:titlePg[^>]*\/>/g, ""));
  const copia = await zip.generateAsync({ type: "arraybuffer" });
  const caixa = document.createElement("div");
  caixa.style.cssText = "position:fixed;left:-20000px;top:0;visibility:hidden;pointer-events:none";
  const estilos = document.createElement("div");
  document.body.append(caixa, estilos);
  try {
    await renderAsync(copia, caixa, estilos, { inWrapper: false, breakPages: true, renderHeaders: true, renderFooters: true, renderFootnotes: false, renderChanges: false, renderComments: false, useBase64URL: true });
    const s = caixa.querySelector("section");
    return { cabecalho: converterArea(s?.querySelector("header")), rodape: converterArea(s?.querySelector("footer")) };
  } finally {
    caixa.remove();
    estilos.remove();
  }
}

function converterArea(el) {
  if (!el) return "";
  const copia = el.cloneNode(true);
  inlinarEstilos(el, copia);
  return limpar(copia);
}

export async function docxParaHtml(buffer) {
  const { renderAsync } = await import("docx-preview");
  const caixa = document.createElement("div");
  caixa.style.cssText = "position:fixed;left:-20000px;top:0;visibility:hidden;pointer-events:none";
  const estilos = document.createElement("div");
  document.body.append(caixa, estilos);
  try {
    await renderAsync(buffer, caixa, estilos, {
      inWrapper: false, ignoreWidth: false, ignoreHeight: false, breakPages: true, ignoreLastRenderedPageBreak: true,
      renderHeaders: true, renderFooters: true, renderFootnotes: true, renderChanges: true, renderComments: true, renderEndnotes: false, useBase64URL: true, experimental: true,
    });
    // O docx-preview calcula a largura das tabulações 0,5 s depois de desenhar.
    if (caixa.querySelector("[class$='-tab-stop']")) await new Promise((ok) => setTimeout(ok, 650));
    const secoes = [...caixa.querySelectorAll("section")];
    await prepararRevisoes(buffer, caixa);
    const comentarios = prepararComentarios(caixa);
    prepararSumario(caixa);
    if (!secoes.length) throw new Error("Documento vazio.");
    // Página e margens da primeira seção
    const st = getComputedStyle(secoes[0]);
    const largura = pxCm(secoes[0].getBoundingClientRect().width || parseFloat(st.width));
    const altura = pxCm(parseFloat(st.minHeight) || parseFloat(st.height) || 0);
    const [a, b] = largura > altura && altura ? [altura, largura] : [largura, altura];
    const tam = Object.entries(PAPEIS).reduce((m, [k, [w, h]]) => (Math.abs(w - a) + Math.abs(h - b) < Math.abs(PAPEIS[m][0] - a) + Math.abs(PAPEIS[m][1] - b) ? k : m), "A4");
    const pagina = {
      tam, orient: largura > altura && altura ? "paisagem" : "retrato",
      margens: { sup: pxCm(st.paddingTop) || 2.5, inf: pxCm(st.paddingBottom) || 2.5, esq: pxCm(st.paddingLeft) || 3, dir: pxCm(st.paddingRight) || 2 },
    };
    // Notas de rodapé: chamadas no texto e notas (lista no fim da seção), na mesma ordem.
    const notas = {};
    for (const s of secoes) {
      const chamadas = [...s.querySelectorAll("article [class*='footnotereference']")].filter((x) => x.querySelector("sup") || x.textContent.trim());
      const itens = [...s.querySelectorAll(":scope > ol > li")];
      chamadas.forEach((ch, i) => {
        if (!itens[i]) return;
        const id = "n" + Math.random().toString(36).slice(2, 9);
        ch.dataset.notaId = id;
        notas[id] = converterArea(itens[i]);
      });
    }
    const corpos = secoes.map((s) => converterArea(s.querySelector("article")));
    // Marcadores de início/fim de comentário viram trechos destacados.
    const juntar = document.createElement("div");
    // Páginas viram quebras de página; mudança de orientação vira quebra de seção.
    const orientDe = (sec) => {
      const cs = getComputedStyle(sec);
      const wd = sec.getBoundingClientRect().width || parseFloat(cs.width);
      const ht = parseFloat(cs.minHeight) || parseFloat(cs.height) || 0;
      return ht && wd > ht ? "paisagem" : "retrato";
    };
    let htmlJunto = "";
    let orientAnt = orientDe(secoes[0]);
    corpos.forEach((c, i) => {
      const o = orientDe(secoes[i]);
      if (i > 0 && o !== orientAnt) htmlJunto += `<hr class="secao" data-orient="${o}">`;
      else if (i > 0 && c && htmlJunto) htmlJunto += '<hr class="quebra">';
      orientAnt = o;
      htmlJunto += c || "";
    });
    juntar.innerHTML = htmlJunto;
    for (const ini of [...juntar.querySelectorAll("[data-com-ini]")]) {
      const id = ini.dataset.comIni;
      const fim = juntar.querySelector(`[data-com-fim="${id}"]`);
      if (fim) {
        const r = document.createRange();
        r.setStartAfter(ini);
        r.setEndBefore(fim);
        envolverIntervalo(r, () => { const sp = document.createElement("span"); sp.className = "coment"; sp.dataset.c = id; return sp; });
      }
    }
    juntar.querySelectorAll("[data-com-ini],[data-com-fim]").forEach((x) => x.remove());
    juntar.querySelectorAll("[data-sumario]").forEach((x) => { x.className = "sumario"; x.setAttribute("contenteditable", "false"); x.innerHTML = ""; });
    for (const id of Object.keys(comentarios)) if (!juntar.querySelector(`[data-c="${id}"]`)) delete comentarios[id];
    let cabecalho = converterArea(secoes[0].querySelector("header"));
    let rodape = converterArea(secoes[0].querySelector("footer"));
    // Primeira página diferente (w:titlePg): o que apareceu é o da 1ª página; o das demais
    // vem de uma segunda renderização sem a marca.
    let extra = {};
    const demais = await cabecalhosSemTitlePg(buffer, renderAsync);
    if (demais) {
      extra = { primeiraDiferente: true, cabecalho1: cabecalho, rodape1: rodape };
      ({ cabecalho, rodape } = demais);
    }
    // Número de página no rodapé (campo PAGE): o docx-preview não mostra; vira a opção do editor.
    const numPag = await numeroDePagina(buffer);
    if (numPag) {
      const so = (h) => { const t = document.createElement("div"); t.innerHTML = h; return t.textContent.trim() || t.querySelector("img") ? h : ""; };
      rodape = so(rodape);
      if (extra.rodape1 !== undefined) extra.rodape1 = so(extra.rodape1);
    }
    return { html: juntar.innerHTML || "<p><br></p>", cabecalho, rodape, ...extra, numPag, pagina, notas, comentarios };
  } finally {
    caixa.remove();
    estilos.remove();
  }
}

// O docx-preview mostra <ins>/<del> mas não diz quem alterou: lemos os autores no XML,
// na mesma ordem em que aparecem.
async function prepararRevisoes(buffer, caixa) {
  try {
    const JSZip = (await import("jszip")).default;
    const xml = await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("string");
    if (!xml) return;
    const x = new DOMParser().parseFromString(xml, "application/xml");
    const lista = (tag) => [...x.getElementsByTagName(tag)].filter((e) => e.parentNode?.nodeName !== "w:rPr" && e.getElementsByTagName("w:r").length)
      .map((e) => ({ autor: e.getAttribute("w:author") || "", data: e.getAttribute("w:date") || "" }));
    for (const [tag, dados] of [["ins", lista("w:ins")], ["del", lista("w:del")]]) {
      [...caixa.querySelectorAll(`article ${tag}`)].forEach((el, i) => {
        if (dados[i]) { el.dataset.autor = dados[i].autor; el.dataset.data = dados[i].data; }
      });
    }
  } catch { /* sem autores: ficam como "Autor desconhecido" */ }
}

// Comentários do docx-preview: marcadores <!--start/end of comment #N--> e um balão com
// autor, data e texto. Viram { id: { autor, data, texto } } e marcadores no texto.
function prepararComentarios(caixa) {
  const comentarios = {};
  const ids = {};
  for (const balao of [...caixa.querySelectorAll(".docx-comment-popover")]) {
    const suporte = balao.parentElement;
    let n = null, data = "";
    for (const ch of suporte.childNodes) {
      if (ch.nodeType !== 8) continue;
      const m = /comment #(\d+)(?: by .*? on (\S+))?/.exec(ch.data);
      if (m) { n = m[1]; data = m[2] || ""; }
    }
    if (n === null) continue;
    const id = "c" + Math.random().toString(36).slice(2, 9);
    ids[n] = id;
    comentarios[id] = {
      autor: balao.querySelector(".docx-comment-author")?.textContent || "Autor desconhecido",
      data: data || new Date().toISOString(),
      texto: [...balao.querySelectorAll("p")].map((p) => p.textContent).join("\n"),
    };
    suporte.remove();
  }
  const w = document.createTreeWalker(caixa, NodeFilter.SHOW_COMMENT);
  const marcas = [];
  for (let c = w.nextNode(); c; c = w.nextNode()) marcas.push(c);
  for (const c of marcas) {
    const m = /^(start|end) of comment #(\d+)/.exec(c.data.trim());
    if (!m || !ids[m[2]]) continue;
    const sp = document.createElement("span");
    sp.setAttribute(m[1] === "start" ? "data-com-ini" : "data-com-fim", ids[m[2]]);
    c.replaceWith(sp);
  }
  return comentarios;
}

// Mede a largura do texto (em pontos) com as fontes que o navegador tem, para estimar em
// que página cada título cai.
async function medidorDeTexto(html, est) {
  const familias = new Set(Object.values(est).map((e) => e.fonte));
  for (const m of String(html).matchAll(/font-family:\s*(?:&quot;|["'])?([^;"'&]+)/g)) familias.add(m[1].trim());
  for (const m of String(html).matchAll(/face="([^"]+)"/g)) familias.add(m[1]);
  await Promise.all([...familias].flatMap((f) => [`11pt "${f}"`, `bold 11pt "${f}"`, `italic 11pt "${f}"`].map((x) => document.fonts.load(x).catch(() => null))));
  const ctx = document.createElement("canvas").getContext("2d");
  const cache = new Map();
  return (texto, { name, size, bold, italic }) => {
    const fonte = `${italic ? "italic " : ""}${bold ? "bold " : ""}${(size * 4) / 3}px "${name}", sans-serif`;
    const chave = fonte + "\n" + texto;
    if (!cache.has(chave)) {
      ctx.font = fonte;
      cache.set(chave, ctx.measureText(texto).width * 0.75);
    }
    return cache.get(chave);
  };
}

// Sumário do Word (parágrafos com estilo "Sumário 1/2/3", em inglês "TOC 1/2/3") volta a ser
// um sumário automático do editor, que se atualiza com os títulos.
function prepararSumario(caixa) {
  for (const art of caixa.querySelectorAll("article")) {
    const linhas = [...art.querySelectorAll(":scope > p")].filter((p) => /(?:^|\s)docx_(?:toc|sum[aá]?rio)[1-9](?:\s|$)/i.test(p.className));
    if (!linhas.length) continue;
    // O título "Sumário" logo antes já faz parte do sumário do editor.
    const antes = linhas[0].previousElementSibling;
    if (antes && /^\s*sum[aá]rio\s*$/i.test(antes.textContent)) antes.remove();
    const marca = document.createElement("div");
    marca.setAttribute("data-sumario", "1");
    linhas[0].before(marca);
    linhas.forEach((p) => p.remove());
  }
}
