import { useEffect, useRef, useState } from "react";

const GOLD = "#9a6e1e";
const CHAVE = "editorTexto.docs";

// localStorage pode falhar (aba anônima, bloqueio do navegador): o editor segue sem ele.
function ler(chave, padrao) {
  try { const v = localStorage.getItem(chave); return v === null ? padrao : JSON.parse(v); } catch { return padrao; }
}
function gravar(chave, valor) {
  try { localStorage.setItem(chave, JSON.stringify(valor)); } catch { /* ignora */ }
}

function novoDoc(nome = "Documento sem título", html = "<p><br></p>") {
  return { id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), nome, html, atualizado: Date.now() };
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

const ALINHAMENTO = { left: "left", center: "center", right: "right", justify: "both", start: "left", end: "right" };
const TITULOS = { H1: "HEADING_1", H2: "HEADING_2", H3: "HEADING_3" };

function corHex(css) {
  if (!css) return undefined;
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(css);
  if (m) return [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("");
  const h = /^#([0-9a-f]{6})$/i.exec(css.trim());
  return h ? h[1] : undefined;
}

// Tamanhos do execCommand("fontSize") 1..7 em pontos.
const FONT_TAMANHO = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 };

async function htmlParaDocx(html, nome) {
  const d = await import("docx");
  const raiz = new DOMParser().parseFromString(`<div>${html}</div>`, "text/html").body.firstChild;

  function runs(no, estilo) {
    const out = [];
    for (const filho of no.childNodes) {
      if (filho.nodeType === 3) {
        if (filho.textContent) out.push(new d.TextRun({ text: filho.textContent, ...estilo }));
        continue;
      }
      if (filho.nodeType !== 1) continue;
      const tag = filho.tagName;
      if (tag === "BR") { out.push(new d.TextRun({ break: 1 })); continue; }
      const s = { ...estilo };
      if (tag === "B" || tag === "STRONG") s.bold = true;
      if (tag === "I" || tag === "EM") s.italics = true;
      if (tag === "U") s.underline = {};
      if (tag === "S" || tag === "STRIKE") s.strike = true;
      if (tag === "FONT") {
        if (filho.getAttribute("size")) s.size = (FONT_TAMANHO[filho.getAttribute("size")] || 12) * 2;
        if (filho.getAttribute("color")) s.color = corHex(filho.getAttribute("color"));
        if (filho.getAttribute("face")) s.font = filho.getAttribute("face");
      }
      const st = filho.style;
      if (st.fontWeight === "bold" || Number(st.fontWeight) >= 600) s.bold = true;
      if (st.fontStyle === "italic") s.italics = true;
      if (st.textDecoration.includes("underline")) s.underline = {};
      if (st.textDecoration.includes("line-through")) s.strike = true;
      if (st.color) s.color = corHex(st.color);
      if (st.backgroundColor && corHex(st.backgroundColor)) s.shading = { type: d.ShadingType.CLEAR, fill: corHex(st.backgroundColor) };
      if (st.fontFamily) s.font = st.fontFamily.split(",")[0].replace(/["']/g, "").trim();
      const pt = /([\d.]+)pt/.exec(st.fontSize) || /([\d.]+)px/.exec(st.fontSize);
      if (pt) s.size = Math.round(parseFloat(pt[1]) * (st.fontSize.endsWith("px") ? 0.75 : 1) * 2);
      out.push(...runs(filho, s));
    }
    return out;
  }

  function paragrafo(el, extra = {}) {
    const alin = ALINHAMENTO[el.style?.textAlign || el.getAttribute?.("align") || ""];
    return new d.Paragraph({
      children: runs(el, {}),
      heading: TITULOS[el.tagName] ? d.HeadingLevel[TITULOS[el.tagName]] : undefined,
      alignment: alin ? d.AlignmentType[{ left: "LEFT", center: "CENTER", right: "RIGHT", both: "JUSTIFIED" }[alin]] : undefined,
      ...extra,
    });
  }

  function tabela(el) {
    const linhas = [...el.querySelectorAll("tr")].map((tr) => new d.TableRow({
      children: [...tr.children].map((td) => new d.TableCell({ children: blocos(td).length ? blocos(td) : [new d.Paragraph("")] })),
    }));
    return new d.Table({ rows: linhas, width: { size: 100, type: d.WidthType.PERCENTAGE } });
  }

  function blocos(no) {
    const out = [];
    let soltos = null; // texto/inline fora de parágrafo
    const fecharSoltos = () => { if (soltos) { out.push(paragrafo(soltos)); soltos = null; } };
    for (const filho of no.childNodes) {
      const tag = filho.nodeType === 1 ? filho.tagName : null;
      if (tag && ["P", "DIV", "H1", "H2", "H3", "BLOCKQUOTE"].includes(tag)) {
        fecharSoltos();
        if ([...filho.children].some((c) => ["P", "DIV", "UL", "OL", "TABLE", "H1", "H2", "H3"].includes(c.tagName))) out.push(...blocos(filho));
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
      } else {
        if (!soltos) soltos = document.createElement("p");
        soltos.appendChild(filho.cloneNode(true));
      }
    }
    fecharSoltos();
    return out;
  }

  const doc = new d.Document({
    numbering: { config: [{ reference: "numerada", levels: [{ level: 0, format: d.LevelFormat.DECIMAL, text: "%1.", alignment: d.AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    styles: { default: { document: { run: { font: "Times New Roman", size: 24 } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 } } },
      children: blocos(raiz),
    }],
  });
  baixar(await d.Packer.toBlob(doc), `${nome}.docx`);
}

// ---------- Componente ----------

const FONTES = ["Times New Roman", "Arial", "Calibri", "Georgia", "Verdana", "Courier New"];
const TAMANHOS = [["1", "8"], ["2", "10"], ["3", "12"], ["4", "14"], ["5", "18"], ["6", "24"], ["7", "36"]];

const sBarra = { display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center", padding: 8, background: "#f4f1ea", border: "1px solid rgba(0,0,0,0.12)", borderRadius: "10px 10px 0 0", position: "sticky", top: 0, zIndex: 2 };
const sBotao = (ativo) => ({ minWidth: 32, height: 30, padding: "0 8px", border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: ativo ? "rgba(154,110,30,0.2)" : "#fff", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif", color: "#2a2a2a" });
const sSelect = { height: 30, border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: "#fff", fontSize: 13, fontFamily: "sans-serif", padding: "0 4px" };
const sSep = { width: 1, height: 22, background: "rgba(0,0,0,0.15)", margin: "0 4px" };

export default function EditorTexto() {
  const [docs, setDocs] = useState(() => {
    const salvos = ler(CHAVE, null);
    return salvos?.length ? salvos : [novoDoc()];
  });
  const [atualId, setAtualId] = useState(() => docs[0].id);
  const [estado, setEstado] = useState({});
  const [msg, setMsg] = useState("");
  const editor = useRef(null);
  const arquivo = useRef(null);
  const atual = docs.find((x) => x.id === atualId) || docs[0];

  useEffect(() => { gravar(CHAVE, docs); }, [docs]);

  // Troca de documento: carrega o HTML no editor (fora do React para não perder o cursor).
  useEffect(() => {
    if (editor.current) editor.current.innerHTML = atual.html;
    document.execCommand("styleWithCSS", false, false);
  }, [atualId]);

  useEffect(() => {
    const atualizarEstado = () => {
      if (!editor.current?.contains(document.getSelection()?.anchorNode)) return;
      setEstado({
        bold: document.queryCommandState("bold"),
        italic: document.queryCommandState("italic"),
        underline: document.queryCommandState("underline"),
        strikeThrough: document.queryCommandState("strikeThrough"),
        insertUnorderedList: document.queryCommandState("insertUnorderedList"),
        insertOrderedList: document.queryCommandState("insertOrderedList"),
        bloco: (document.queryCommandValue("formatBlock") || "p").toLowerCase(),
      });
    };
    document.addEventListener("selectionchange", atualizarEstado);
    return () => document.removeEventListener("selectionchange", atualizarEstado);
  }, []);

  function salvarHtml() {
    const html = editor.current.innerHTML;
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, html, atualizado: Date.now() } : x)));
  }

  function cmd(c, v = null) {
    editor.current.focus();
    document.execCommand(c, false, v);
    salvarHtml();
  }

  function renomear(nome) {
    setDocs((ds) => ds.map((x) => (x.id === atualId ? { ...x, nome } : x)));
  }

  function criar(nome, html) {
    const n = novoDoc(nome, html);
    setDocs((ds) => [n, ...ds]);
    setAtualId(n.id);
  }

  function excluir() {
    if (!confirm(`Excluir "${atual.nome}"?`)) return;
    const resto = docs.filter((x) => x.id !== atualId);
    const lista = resto.length ? resto : [novoDoc()];
    setDocs(lista);
    setAtualId(lista[0].id);
  }

  function inserirTabela() {
    const l = parseInt(prompt("Número de linhas:", "3"), 10);
    const c = parseInt(prompt("Número de colunas:", "3"), 10);
    if (!(l > 0 && c > 0)) return;
    const td = '<td style="border:1px solid #888;padding:4px 6px;min-width:40px"><br></td>';
    const tr = `<tr>${td.repeat(c)}</tr>`;
    cmd("insertHTML", `<table style="border-collapse:collapse;width:100%">${tr.repeat(l)}</table><p><br></p>`);
  }

  function inserirLink() {
    const url = prompt("Endereço do link:", "https://");
    if (url) cmd("createLink", url);
  }

  async function abrirArquivo(e) {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    const nome = f.name.replace(/\.[^.]+$/, "");
    try {
      if (/\.docx$/i.test(f.name)) {
        setMsg("Abrindo arquivo...");
        const mammoth = await import("mammoth");
        const r = await mammoth.convertToHtml({ arrayBuffer: await f.arrayBuffer() });
        criar(nome, r.value || "<p><br></p>");
      } else if (/\.(txt|md)$/i.test(f.name)) {
        const t = await f.text();
        const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
        criar(nome, t.split(/\r?\n/).map((l) => `<p>${esc(l) || "<br>"}</p>`).join(""));
      } else if (/\.html?$/i.test(f.name)) {
        const corpo = new DOMParser().parseFromString(await f.text(), "text/html").body;
        corpo.querySelectorAll("script,style,iframe,object,embed,link,meta").forEach((n) => n.remove());
        corpo.querySelectorAll("*").forEach((n) => [...n.attributes].forEach((a) => { if (/^on/i.test(a.name) || /^\s*javascript:/i.test(a.value)) n.removeAttribute(a.name); }));
        criar(nome, corpo.innerHTML);
      } else {
        throw new Error("Formato não suportado. Use .docx, .txt ou .html.");
      }
      setMsg("");
    } catch (err) {
      setMsg(err.message || "Não foi possível abrir o arquivo.");
    }
  }

  async function salvarDocx() {
    salvarHtml();
    setMsg("Gerando .docx...");
    try {
      await htmlParaDocx(editor.current.innerHTML, atual.nome || "documento");
      setMsg("");
    } catch (err) {
      setMsg("Erro ao gerar .docx: " + err.message);
    }
  }

  function imprimir() {
    const w = window.open("", "_blank");
    if (!w) return setMsg("Permita pop-ups para imprimir ou salvar em PDF.");
    w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${atual.nome.replace(/</g, "&lt;")}</title>
      <style>@page{size:A4;margin:3cm 2cm 2cm 3cm}body{font-family:'Times New Roman',serif;font-size:12pt;line-height:1.5}table{border-collapse:collapse}hr{page-break-after:always;border:0}</style>
      </head><body>${editor.current.innerHTML}</body></html>`);
    w.document.close();
    w.focus();
    w.print();
  }

  function teclas(e) {
    if (!(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    if (k === "s") { e.preventDefault(); salvarDocx(); }
    if (k === "p") { e.preventDefault(); imprimir(); }
  }

  const palavras = (atual.html.replace(/<[^>]+>/g, " ").match(/\S+/g) || []).length;

  return (
    <div style={{ fontFamily: "sans-serif" }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
        <select value={atualId} onChange={(e) => { salvarHtml(); setAtualId(e.target.value); }} style={{ ...sSelect, maxWidth: 220 }} title="Documentos salvos neste navegador">
          {docs.map((x) => <option key={x.id} value={x.id}>{x.nome || "(sem nome)"}</option>)}
        </select>
        <input value={atual.nome} onChange={(e) => renomear(e.target.value)} style={{ ...sSelect, flex: "1 1 180px", padding: "0 8px" }} placeholder="Nome do documento" />
        <button style={sBotao()} onClick={() => criar()}>＋ Novo</button>
        <button style={sBotao()} onClick={() => arquivo.current.click()}>📂 Abrir</button>
        <button style={sBotao()} onClick={salvarDocx}>💾 Word (.docx)</button>
        <button style={sBotao()} onClick={imprimir}>🖨️ PDF / Imprimir</button>
        <button style={sBotao()} onClick={excluir} title="Excluir documento">🗑️</button>
        <input ref={arquivo} type="file" accept=".docx,.txt,.md,.html,.htm" onChange={abrirArquivo} style={{ display: "none" }} />
      </div>
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
        <select style={sSelect} defaultValue="" onChange={(e) => { cmd("fontName", e.target.value); e.target.value = ""; }} title="Fonte">
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
        <button style={sBotao(estado.insertUnorderedList)} onClick={() => cmd("insertUnorderedList")} title="Marcadores">• Lista</button>
        <button style={sBotao(estado.insertOrderedList)} onClick={() => cmd("insertOrderedList")} title="Numeração">1. Lista</button>
        <button style={sBotao()} onClick={() => cmd("outdent")} title="Diminuir recuo">⇤</button>
        <button style={sBotao()} onClick={() => cmd("indent")} title="Aumentar recuo">⇥</button>
        <span style={sSep} />
        <button style={sBotao()} onClick={inserirTabela} title="Inserir tabela">▦ Tabela</button>
        <button style={sBotao()} onClick={inserirLink} title="Inserir link">🔗</button>
        <button style={sBotao()} onClick={() => cmd("insertHTML", '<hr style="border:0;border-top:1px dashed #bbb"><p><br></p>')} title="Quebra de página">⤓ Página</button>
        <button style={sBotao()} onClick={() => cmd("removeFormat")} title="Limpar formatação">⌫ Formato</button>
      </div>

      <div style={{ background: "#e9e6df", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0, borderRadius: "0 0 10px 10px", padding: "24px 12px", overflowX: "auto" }}>
        <div
          ref={editor}
          contentEditable
          suppressContentEditableWarning
          spellCheck
          lang="pt-BR"
          onInput={salvarHtml}
          onKeyDown={teclas}
          style={{ background: "#fff", width: "100%", maxWidth: "21cm", minHeight: "29.7cm", margin: "0 auto", padding: "3cm 2cm 2cm 3cm", boxSizing: "border-box", boxShadow: "0 2px 12px rgba(0,0,0,0.15)", fontFamily: "'Times New Roman', serif", fontSize: "12pt", lineHeight: 1.5, color: "#111", outline: "none" }}
        />
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#666", marginTop: 6 }}>
        <span>{palavras} palavra{palavras === 1 ? "" : "s"} · salvo automaticamente neste navegador</span>
        <span>Ctrl+S salva .docx · Ctrl+P imprime</span>
      </div>
    </div>
  );
}
