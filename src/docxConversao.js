// Conversão entre o HTML do editor de Texto e arquivos .docx do Word.

export const CM = 37.7953; // pixels por centímetro na tela
export const PAPEIS = { A4: [21, 29.7], Oficio: [21.6, 33], Carta: [21.59, 27.94] };
export const PAGINA_PADRAO = { tam: "A4", orient: "retrato", margens: { sup: 3, inf: 2, esq: 3, dir: 2 } };
export const FONTE_PADRAO = "Times New Roman";
export const TAMANHO_PADRAO = 12; // pt

export function dimensoes(pagina = PAGINA_PADRAO) {
  const [w, h] = PAPEIS[pagina.tam] || PAPEIS.A4;
  return pagina.orient === "paisagem" ? [h, w] : [w, h];
}

function baixar(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nome;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------- HTML do editor → .docx ----------

const TWIP_CM = 566.93;
const ALINHAMENTO = { left: "LEFT", start: "LEFT", center: "CENTER", right: "RIGHT", end: "RIGHT", justify: "JUSTIFIED" };
const TITULOS = { H1: "HEADING_1", H2: "HEADING_2", H3: "HEADING_3" };
const FONT_TAMANHO = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 }; // execCommand("fontSize")

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

export async function htmlParaDocx(doc, { html, cabecalho, rodape }) {
  const d = await import("docx");
  const parse = (h) => new DOMParser().parseFromString(`<div>${h || ""}</div>`, "text/html").body.firstChild;
  const pagina = doc.pagina || PAGINA_PADRAO;
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
    if (corHex(st.backgroundColor)) s.shading = { type: d.ShadingType.CLEAR, fill: corHex(st.backgroundColor) };
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
    return new d.ImageRun({ type: dados.tipo, data: dados.bytes, transformation: { width: Math.round(w), height: Math.round(h) } });
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
        const partes = filho.textContent.split("\t");
        partes.forEach((p, i) => {
          if (i > 0) out.push(new d.TextRun({ children: [new d.Tab()], ...estilo }));
          if (p) out.push(new d.TextRun({ text: p.replace(/\n/g, " "), ...estilo }));
        });
        continue;
      }
      if (filho.nodeType !== 1) continue;
      const tag = filho.tagName;
      if (tag === "BR") {
        // O <br> no fim do parágrafo só marca linha vazia no HTML; no Word viraria uma linha a mais.
        if (!temConteudoDepois(filho)) continue;
        out.push(new d.TextRun({ break: 1, ...estilo }));
        continue;
      }
      if (tag === "IMG") { const im = imagem(filho); if (im) out.push(im); continue; }
      if (filho.dataset?.campo === "pagina") { out.push(new d.TextRun({ children: [d.PageNumber.CURRENT], ...estilo })); continue; }
      if (filho.dataset?.campo === "paginas") { out.push(new d.TextRun({ children: [d.PageNumber.TOTAL_PAGES], ...estilo })); continue; }
      if (tag === "A" && filho.getAttribute("href")) {
        out.push(new d.ExternalHyperlink({ link: filho.getAttribute("href"), children: runs(filho, { ...estiloRun(filho, estilo), style: "Hyperlink" }) }));
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
    if (Object.keys(bordas).length) p.border = bordas;
    return p;
  }

  function paragrafo(el, extra = {}) {
    const base = estiloRun(el, {});
    // Fonte da marca de parágrafo: define a altura de linhas em branco no Word.
    const marca = {};
    if (base.font) marca.font = base.font;
    if (base.size) marca.size = base.size;
    return new d.Paragraph({
      children: runs(el, base),
      run: Object.keys(marca).length ? marca : undefined,
      heading: TITULOS[el.tagName] ? d.HeadingLevel[TITULOS[el.tagName]] : undefined,
      ...propsParagrafo(el),
      ...extra,
    });
  }

  function borda(css) {
    const m = /([\d.]+)px\s+(\w+)\s+(.*)/.exec(css || "");
    if (!m || m[2] === "none" || parseFloat(m[1]) === 0) return { style: d.BorderStyle.NONE, size: 0, color: "FFFFFF" };
    return { style: m[2] === "dashed" ? d.BorderStyle.DASHED : m[2] === "dotted" ? d.BorderStyle.DOTTED : d.BorderStyle.SINGLE, size: Math.max(2, Math.round(parseFloat(m[1]) * 6)), color: corHex(m[3]) || "000000" };
  }

  function tabela(el) {
    const linhas = [...el.rows].map((tr) => new d.TableRow({
      children: [...tr.cells].map((td) => {
        const conteudo = blocos(td);
        const st = td.style;
        const temBorda = st.borderTop || st.borderBottom || st.borderLeft || st.borderRight || st.border;
        return new d.TableCell({
          children: conteudo.length ? conteudo : [new d.Paragraph("")],
          columnSpan: td.colSpan > 1 ? td.colSpan : undefined,
          rowSpan: td.rowSpan > 1 ? td.rowSpan : undefined,
          shading: corHex(st.backgroundColor) ? { type: d.ShadingType.CLEAR, fill: corHex(st.backgroundColor) } : undefined,
          borders: temBorda ? {
            top: borda(st.borderTop || st.border), bottom: borda(st.borderBottom || st.border),
            left: borda(st.borderLeft || st.border), right: borda(st.borderRight || st.border),
          } : undefined,
        });
      }),
    }));
    return new d.Table({ rows: linhas, width: { size: 100, type: d.WidthType.PERCENTAGE } });
  }

  const BLOCO = ["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE"];
  function blocos(no) {
    const out = [];
    let soltos = null; // texto e elementos de linha fora de parágrafo
    const fecharSoltos = () => { if (soltos) { out.push(paragrafo(soltos)); soltos = null; } };
    for (const filho of no.childNodes) {
      const tag = filho.nodeType === 1 ? filho.tagName : null;
      if (tag && BLOCO.includes(tag)) {
        fecharSoltos();
        if ([...filho.children].some((c) => [...BLOCO, "UL", "OL", "TABLE"].includes(c.tagName))) out.push(...blocos(filho));
        else out.push(paragrafo(filho, tag === "BLOCKQUOTE" ? { indent: { left: 720 } } : {}));
      } else if (tag === "UL" || tag === "OL") {
        fecharSoltos();
        for (const li of filho.querySelectorAll(":scope > li")) {
          out.push(paragrafo(li, tag === "UL" ? { bullet: { level: 0 } } : { numbering: { reference: "numerada", level: 0 } }));
        }
      } else if (tag === "TABLE") {
        fecharSoltos();
        out.push(tabela(filho));
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
  const rodapeFilhos = blocos(parse(rodape));
  if (doc.numPag) {
    const pagTxt = doc.numPag === "xdey"
      ? [new d.TextRun({ children: ["Página ", d.PageNumber.CURRENT, " de ", d.PageNumber.TOTAL_PAGES] })]
      : [new d.TextRun({ children: [d.PageNumber.CURRENT] })];
    rodapeFilhos.push(new d.Paragraph({ alignment: doc.numPag === "centro" || doc.numPag === "xdey" ? d.AlignmentType.CENTER : d.AlignmentType.RIGHT, children: pagTxt }));
  }
  const cabFilhos = blocos(parse(cabecalho));

  const titulo = (id, nome, pt) => ({
    id, name: nome, basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { size: pt * 2, bold: true, color: "000000", font: FONTE_PADRAO },
    paragraph: { spacing: { before: 240, after: 120 }, keepNext: true },
  });
  const documento = new d.Document({
    numbering: { config: [{ reference: "numerada", levels: [{ level: 0, format: d.LevelFormat.DECIMAL, text: "%1.", alignment: d.AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    styles: {
      default: { document: { run: { font: FONTE_PADRAO, size: TAMANHO_PADRAO * 2 }, paragraph: { spacing: { line: 360, lineRule: d.LineRuleType.AUTO, after: 120 } } } },
      paragraphStyles: [titulo("Heading1", "Heading 1", 16), titulo("Heading2", "Heading 2", 14), titulo("Heading3", "Heading 3", 12)],
    },
    sections: [{
      properties: {
        page: {
          size: { width: Math.round(w * TWIP_CM), height: Math.round(h * TWIP_CM), orientation: pagina.orient === "paisagem" ? d.PageOrientation.LANDSCAPE : d.PageOrientation.PORTRAIT },
          margin: { top: Math.round(m.sup * TWIP_CM), bottom: Math.round(m.inf * TWIP_CM), left: Math.round(m.esq * TWIP_CM), right: Math.round(m.dir * TWIP_CM), header: Math.round(1.25 * TWIP_CM), footer: Math.round(1.25 * TWIP_CM) },
        },
      },
      headers: cabFilhos.length ? { default: new d.Header({ children: cabFilhos }) } : undefined,
      footers: rodapeFilhos.length ? { default: new d.Footer({ children: rodapeFilhos }) } : undefined,
      children: (() => { const b = blocos(parse(html)); return b.length ? b : [new d.Paragraph("")]; })(),
    }],
  });
  baixar(await d.Packer.toBlob(documento), `${doc.nome || "documento"}.docx`);
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
    const st = getComputedStyle(o);
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
        for (const lado of ["Top", "Bottom"]) {
          const estilo = st[`border${lado}Style`];
          if (estilo && estilo !== "none" && parseFloat(st[`border${lado}Width`]) > 0) decl.push(`border-${lado.toLowerCase()}:${st[`border${lado}Width`]} ${estilo} ${st[`border${lado}Color`]}`, `padding-${lado.toLowerCase()}:2pt`);
        }
        if (parseFloat(st.marginLeft) > 0) decl.push(`margin-left:${pxCm(st.marginLeft)}cm`);
        else if (parseFloat(st.paddingLeft) > 0) decl.push(`margin-left:${pxCm(st.paddingLeft)}cm`);
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
    for (const a of [...c.attributes]) if (!["href", "src", "colspan", "rowspan"].includes(a.name)) c.removeAttribute(a.name);
    if (tag === "IMG") {
      c.setAttribute("width", Math.round(o.getBoundingClientRect().width) || o.naturalWidth || 300);
      c.setAttribute("height", Math.round(o.getBoundingClientRect().height) || o.naturalHeight || 200);
    }
    if (tag === "TABLE") decl.push("border-collapse:collapse", "width:100%");
    if (decl.length) c.setAttribute("style", decl.join(";"));
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
      if (numero) c.insertBefore(document.createTextNode(numero + "\t"), c.firstChild);
    }
    if (o.classList.contains?.("docx-tab-stop") || /tab-stop/.test(o.className)) c.replaceWith(document.createTextNode("\t"));
  });
}

function limpar(raiz) {
  // Tira embrulhos vazios (span sem estilo) e elementos que o editor não usa.
  raiz.querySelectorAll("script,style,link,meta").forEach((n) => n.remove());
  raiz.querySelectorAll("p div, h1 div, h2 div, h3 div, h4 div, li div").forEach((n) => {
    const span = document.createElement("span");
    span.append(...n.childNodes);
    n.replaceWith(span);
  });
  raiz.querySelectorAll("span:not([style]),font:not([style])").forEach((n) => n.replaceWith(...n.childNodes));
  raiz.querySelectorAll("p").forEach((p) => { if (!p.textContent && !p.querySelector("img,br")) p.innerHTML = "<br>"; });
  return raiz.innerHTML;
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
      renderHeaders: true, renderFooters: true, renderFootnotes: false, renderEndnotes: false, useBase64URL: true, experimental: true,
    });
    const secoes = [...caixa.querySelectorAll("section")];
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
    const corpos = secoes.map((s) => converterArea(s.querySelector("article")));
    const cabecalho = converterArea(secoes[0].querySelector("header"));
    const rodape = converterArea(secoes[0].querySelector("footer"));
    return { html: corpos.filter(Boolean).join('<hr class="quebra">') || "<p><br></p>", cabecalho, rodape, pagina };
  } finally {
    caixa.remove();
    estilos.remove();
  }
}
