import { Fragment, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { flushSync } from "react-dom";
import {
  analisar, avaliar, valorFinal, interpretarEntrada, exibir, paraTexto, Erro, ERROS, erroEmIngles, erroDoExcel,
  chave, lerChave, nomeColuna, indiceColuna, formulaParaIngles, formulaParaPortugues, deslocarFormula, transformarRefs,
  dataParaSerie,
} from "./formulas.js";

const GOLD = "#9a6e1e";
const CHAVE = "planilha.pastas";
const LARG_PADRAO = 100;
const ALT_LINHA = 24;

function ler(k, padrao) {
  try { const v = localStorage.getItem(k); return v === null ? padrao : JSON.parse(v); } catch { return padrao; }
}
function gravar(k, valor) {
  try { localStorage.setItem(k, JSON.stringify(valor)); return true; } catch { return false; }
}
function baixar(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nome;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const novaAba = (nome = "Planilha1") => ({ nome, cel: {}, larg: {}, alt: {}, mescl: [], linhas: 100, colunas: 26 });
const novaPasta = (nome = "Pasta sem título", planilhas = [novaAba()]) => ({
  id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), nome, planilhas, atualizado: Date.now(),
});

// ---------- Cálculo ----------

const arvores = new Map();
function compilar(f) {
  if (!arvores.has(f)) {
    let a;
    try { a = analisar(f); } catch (e) { a = e instanceof Erro ? e : ERROS.valor; }
    arvores.set(f, a);
  }
  return arvores.get(f);
}

// Valor de uma célula sem fórmula. Apóstrofo inicial força texto, como no Excel.
const valorDigitado = (e) => (e[0] === "'" ? e.slice(1) : interpretarEntrada(e).v);

function criarCalculadora(planilhas) {
  const cache = new Map();
  const visitando = new Set();
  const porNome = new Map(planilhas.map((p, i) => [p.nome.toUpperCase(), i]));
  function valor(a, r, c) {
    const k = a + ":" + r + ":" + c;
    if (cache.has(k)) return cache.get(k);
    const e = planilhas[a]?.cel[chave(r, c)]?.e;
    if (!e) return null;
    if (visitando.has(k)) return ERROS.ciclo;
    visitando.add(k);
    let v;
    if (e[0] === "=" && e.length > 1) {
      const arv = compilar(e.slice(1));
      v = arv instanceof Erro ? arv : valorFinal(avaliar(arv, {
        celula: (p, rr, cc) => {
          if (rr < 0 || cc < 0) return ERROS.ref;
          if (!p) return valor(a, rr, cc);
          const i = porNome.get(p.toUpperCase());
          return i === undefined ? ERROS.ref : valor(i, rr, cc);
        },
      }));
    } else v = valorDigitado(e);
    visitando.delete(k);
    cache.set(k, v);
    return v;
  }
  return valor;
}

// Fórmulas de data mostram data mesmo sem formato escolhido.
const ehFormulaData = (e) => /^=\s*(HOJE|TODAY|DATA|DATE|AGORA|NOW)\s*\(/i.test(e || "");

// ---------- Inserir/excluir linhas e colunas ----------

// eixo "r" (linha) ou "c" (coluna); delta +1 insere antes de pos, -1 exclui pos.
function reestruturar(planilhas, abaIdx, eixo, pos, delta) {
  const nomeAba = planilhas[abaIdx].nome.toUpperCase();
  const mover = (t, papel) => {
    if (t[eixo] < pos) return;
    if (delta > 0) { t[eixo] += 1; return; }
    if (t[eixo] > pos) t[eixo] -= 1;
    else if (papel === "fim") t[eixo] -= 1;
    else if (papel !== "ini") t[eixo] = -1;
  };
  return planilhas.map((p, i) => {
    const cel = {};
    for (const [k, v] of Object.entries(p.cel)) {
      let nova = v;
      if (v.e?.[0] === "=") {
        const e = "=" + transformarRefs(v.e.slice(1), (t, papel) => {
          const alvo = t.planilha ? t.planilha.toUpperCase() === nomeAba : i === abaIdx;
          if (alvo) mover(t, papel);
        });
        if (e !== v.e) nova = { ...v, e };
      }
      if (i === abaIdx) {
        const rc = lerChave(k);
        if (rc[eixo] === pos && delta < 0) continue;
        if (rc[eixo] >= pos) rc[eixo] += delta;
        cel[chave(rc.r, rc.c)] = nova;
      } else cel[k] = nova;
    }
    if (i !== abaIdx) return { ...p, cel };
    const larg = {};
    for (const [c, w] of Object.entries(p.larg || {})) {
      const n = Number(c);
      if (eixo === "c" && n === pos && delta < 0) continue;
      larg[eixo === "c" && n >= pos ? n + delta : n] = w;
    }
    const alt = {};
    for (const [r, h] of Object.entries(p.alt || {})) {
      const n = Number(r);
      if (eixo === "r" && n === pos && delta < 0) continue;
      alt[eixo === "r" && n >= pos ? n + delta : n] = h;
    }
    const lo = eixo === "r" ? "r1" : "c1", hi = eixo === "r" ? "r2" : "c2";
    const mescl = (p.mescl || []).map((m) => {
      const x = { ...m };
      if (delta > 0) { if (x[lo] >= pos) x[lo]++; if (x[hi] >= pos) x[hi]++; }
      else { if (x[lo] > pos) x[lo]--; if (x[hi] >= pos) x[hi]--; }
      return x;
    }).filter((m) => m.r2 >= m.r1 && m.c2 >= m.c1 && (m.r2 > m.r1 || m.c2 > m.c1));
    return {
      ...p, cel, larg, alt, mescl, filtro: moverFiltro(p.filtro, eixo, pos, delta),
      linhas: eixo === "r" ? Math.max(1, p.linhas + delta) : p.linhas,
      colunas: eixo === "c" ? Math.max(1, p.colunas + delta) : p.colunas,
    };
  });
}

function linhasOcultas(planilha, valor, a) {
  const ocultas = new Set();
  const f = planilha.filtro;
  if (!f) return ocultas;
  const crit = Object.entries(f.excl || {}).filter(([, v]) => v.length).map(([c, v]) => [Number(c), new Set(v)]);
  if (!crit.length) return ocultas;
  const fim = fimFiltro(planilha, f);
  for (let r = f.r + 1; r <= fim; r++) {
    if (crit.some(([c, excl]) => excl.has(textoFiltro(planilha, valor, a, r, c)))) ocultas.add(r);
  }
  return ocultas;
}

function moverFiltro(f, eixo, pos, delta) {
  if (!f) return f;
  if (eixo === "r") {
    if (f.r < pos) return f;
    if (f.r === pos && delta < 0) return undefined; // a linha de títulos foi excluída
    return { ...f, r: f.r + delta };
  }
  if (delta < 0 && f.c1 === f.c2 && f.c1 === pos) return undefined;
  const excl = {};
  for (const [c, v] of Object.entries(f.excl || {})) {
    const n = Number(c);
    if (n === pos && delta < 0) continue;
    excl[n >= pos ? n + delta : n] = v;
  }
  const ajusta = (c, fim) => (delta > 0 ? (c >= pos ? c + 1 : c) : c > pos || (c === pos && fim) ? c - 1 : c);
  return { ...f, c1: ajusta(f.c1, false), c2: ajusta(f.c2, true), excl };
}

// ---------- Ordenar e filtrar ----------

// Ordem do Excel: números (e datas), textos, lógicos, erros; vazias sempre por último.
function compararOrdem(a, b, desc) {
  const vazio = (v) => v === null || v === "";
  const grupo = (v) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3);
  if (vazio(a) || vazio(b)) return vazio(a) - vazio(b);
  if (a instanceof Erro || b instanceof Erro) return (a instanceof Erro) - (b instanceof Erro);
  const ga = grupo(a), gb = grupo(b);
  let k = ga !== gb ? ga - gb : ga === 1 ? a.localeCompare(b, "pt-BR", { sensitivity: "base", numeric: true }) : a === b ? 0 : a < b ? -1 : 1;
  return desc ? -k : k;
}

// Texto que identifica o valor no filtro (o mesmo que aparece na célula).
function textoFiltro(planilha, valor, a, r, c) {
  const cel = planilha.cel[chave(r, c)];
  if (!cel) return "";
  return exibir(valor(a, r, c), cel.fmt || (ehFormulaData(cel.e) ? "data" : undefined));
}

// Última linha com dados nas colunas do filtro.
function fimFiltro(planilha, f) {
  let fim = f.r;
  for (const k of Object.keys(planilha.cel)) {
    const { r, c } = lerChave(k);
    if (r > fim && c >= f.c1 && c <= f.c2 && planilha.cel[k].e) fim = r;
  }
  return fim;
}

// Bloco contínuo de dados em volta da célula (a "região atual" do Excel).
function regiaoAtual(planilha, r, c) {
  const tem = (rr, cc) => rr >= 0 && cc >= 0 && rr < planilha.linhas && cc < planilha.colunas && !!planilha.cel[chave(rr, cc)]?.e;
  const linhaTem = (rr, a, b) => { for (let cc = a; cc <= b; cc++) if (tem(rr, cc)) return true; return false; };
  const colunaTem = (cc, a, b) => { for (let rr = a; rr <= b; rr++) if (tem(rr, cc)) return true; return false; };
  let r1 = r, r2 = r, c1 = c, c2 = c;
  for (let mudou = true; mudou;) {
    mudou = false;
    if (linhaTem(r1 - 1, c1 - 1, c2 + 1)) { r1--; mudou = true; }
    if (linhaTem(r2 + 1, c1 - 1, c2 + 1)) { r2++; mudou = true; }
    if (colunaTem(c1 - 1, r1 - 1, r2 + 1)) { c1--; mudou = true; }
    if (colunaTem(c2 + 1, r1 - 1, r2 + 1)) { c2++; mudou = true; }
  }
  return { r1, r2, c1, c2 };
}

// ---------- Alça de preenchimento: séries ----------

const LISTAS_SERIE = [
  ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
  ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"],
  ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"],
  ["domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado"],
  ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"],
];
const mesmaCaixa = (modelo, t) => (modelo === modelo.toUpperCase() ? t.toUpperCase() : modelo[0] === modelo[0].toUpperCase() ? t[0].toUpperCase() + t.slice(1) : t);

// Recebe as células de origem na ordem do preenchimento e devolve o conteúdo das
// próximas qtd células, ou null quando o certo é repetir/copiar a origem.
function gerarSerie(cels, qtd) {
  const brutos = cels.map((c) => c.e || "");
  const n = brutos.length;
  if (brutos.some((e) => e === "" || e[0] === "=")) return null;
  const nums = brutos.map((e) => { if (e[0] === "'") return null; const v = interpretarEntrada(e).v; return typeof v === "number" ? v : null; });
  if (nums.every((x) => x !== null)) {
    const ultima = cels[n - 1];
    if (n === 1 && ultima.fmt !== "data") return null; // número sozinho: copia, como no Excel
    const passo = n === 1 ? 1 : (nums[n - 1] - nums[0]) / (n - 1);
    return Array.from({ length: qtd }, (_, t) => {
      const v = Number((nums[n - 1] + passo * (t + 1)).toPrecision(15));
      if (ultima.fmt === "data") return exibir(v, "data");
      if (ultima.fmt === "pct") return paraTexto(Number((v * 100).toPrecision(15))) + "%";
      return paraTexto(v);
    });
  }
  const partes = brutos.map((e) => /^(.*?)(\d+)(\D*)$/.exec(e));
  if (partes.every(Boolean) && partes.every((x) => x[1] === partes[0][1] && x[3] === partes[0][3])) {
    const ns = partes.map((x) => Number(x[2]));
    const passo = n === 1 ? 1 : (ns[n - 1] - ns[0]) / (n - 1);
    if (Number.isInteger(passo)) {
      const largura = partes[n - 1][2].length;
      return Array.from({ length: qtd }, (_, t) => {
        const v = ns[n - 1] + passo * (t + 1);
        return partes[0][1] + (v < 0 ? String(v) : String(v).padStart(largura, "0")) + partes[0][3];
      });
    }
  }
  for (const lista of LISTAS_SERIE) {
    const idx = brutos.map((e) => lista.indexOf(e.toLowerCase()));
    if (idx.some((i) => i < 0)) continue;
    const passo = n === 1 ? 1 : (idx[n - 1] - idx[0]) / (n - 1);
    if (!Number.isInteger(passo)) continue;
    const L = lista.length;
    return Array.from({ length: qtd }, (_, t) => mesmaCaixa(brutos[n - 1], lista[(((idx[n - 1] + passo * (t + 1)) % L) + L) % L]));
  }
  return null;
}

// Propriedades de uma célula; usado para limpar tudo ao copiar por cima.
const CEL_VAZIA = { e: undefined, b: undefined, i: undefined, al: undefined, fmt: undefined, bg: undefined, cor: undefined, qb: undefined, bd: undefined, tam: undefined, fn: undefined };

// ---------- Excel (.xlsx) ----------

const NUMFMT = { numero: "#,##0.00", moeda: '"R$" #,##0.00', pct: "0.00%", data: "dd/mm/yyyy", inteiro: "#,##0" };
function fmtDoExcel(nf) {
  if (!nf || nf === "General") return undefined;
  if (/R\$|BRL/.test(nf)) return "moeda";
  if (nf.includes("%")) return "pct";
  if (/[dy]/i.test(nf.replace(/"[^"]*"/g, "")) && /m/i.test(nf)) return "data";
  if (/0[.,]0/.test(nf)) return "numero";
  if (/#,##0$/.test(nf)) return "inteiro";
  return undefined;
}
const argb = (hex) => "FF" + hex.replace("#", "").toUpperCase();
const hexDeArgb = (c) => (c?.argb && c.argb.length === 8 ? "#" + c.argb.slice(2) : undefined);

async function exportarXlsx(pasta, valor) {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  pasta.planilhas.forEach((p, a) => {
    const ws = wb.addWorksheet(p.nome.slice(0, 31) || `Planilha${a + 1}`);
    if (p.filtro) {
      const f = p.filtro;
      const fim = fimFiltro(p, f);
      ws.autoFilter = `${chave(f.r, f.c1)}:${chave(fim, f.c2)}`;
      for (const r of linhasOcultas(p, valor, a)) ws.getRow(r + 1).hidden = true;
    }
    for (const [c, w] of Object.entries(p.larg || {})) ws.getColumn(Number(c) + 1).width = Math.round(w / 7);
    for (const [k, cel] of Object.entries(p.cel)) {
      const x = ws.getCell(k);
      const { r, c } = lerChave(k);
      const v = valor(a, r, c);
      const resultado = v instanceof Erro ? { error: erroEmIngles(v) } : v;
      if (cel.e?.[0] === "=" && cel.e.length > 1) x.value = { formula: formulaParaIngles(cel.e.slice(1)), result: resultado ?? undefined };
      else if (cel.e) x.value = v instanceof Erro ? { error: erroEmIngles(v) } : v;
      const fmt = cel.fmt || (ehFormulaData(cel.e) ? "data" : undefined);
      if (fmt && NUMFMT[fmt]) x.numFmt = NUMFMT[fmt];
      if (cel.b || cel.i || cel.cor || cel.tam || cel.fn) {
        x.font = { bold: !!cel.b, italic: !!cel.i, ...(cel.tam ? { size: cel.tam } : {}), ...(cel.fn ? { name: cel.fn } : {}), ...(cel.cor ? { color: { argb: argb(cel.cor) } } : {}) };
      }
      if (cel.al || cel.qb) x.alignment = { ...(cel.al ? { horizontal: cel.al } : {}), ...(cel.qb ? { wrapText: true, vertical: "top" } : {}) };
      if (cel.bd) x.border = Object.fromEntries([...cel.bd].map((l) => [{ t: "top", b: "bottom", l: "left", r: "right" }[l], { style: "thin" }]));
      if (cel.bg) x.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(cel.bg) } };
    }
    for (const m of p.mescl || []) ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    for (const [r, h] of Object.entries(p.alt || {})) ws.getRow(Number(r) + 1).height = Math.round(h * 0.75);
    if (p.cong && (p.cong.r || p.cong.c)) ws.views = [{ state: "frozen", xSplit: p.cong.c, ySplit: p.cong.r }];
  });
  const buf = await wb.xlsx.writeBuffer();
  baixar(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `${pasta.nome || "planilha"}.xlsx`);
}

async function importarXlsx(arquivo) {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await arquivo.arrayBuffer());
  const planilhas = [];
  wb.eachSheet((ws) => {
    const p = novaAba(ws.name);
    let maxR = 0, maxC = 0;
    // includeEmpty: células sem valor também podem ter formatação (bordas, cores)
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      if (row.height > 15.5) p.alt[rn - 1] = Math.round(row.height / 0.75);
      row.eachCell({ includeEmpty: true }, (x, cn) => {
        if (x.isMerged && x.master && x.master.address !== x.address) return; // parte de célula mesclada
        const cel = {};
        const v = x.value;
        let fmt = fmtDoExcel(x.numFmt);
        if (x.formula) cel.e = "=" + formulaParaPortugues(x.formula);
        else if (v instanceof Date) { cel.e = paraTexto(dataParaSerie(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate())); fmt = "data"; }
        else if (typeof v === "number") cel.e = paraTexto(v);
        else if (typeof v === "boolean") cel.e = v ? "VERDADEIRO" : "FALSO";
        else if (typeof v === "string") cel.e = /^[=+-]|^\d|^R\$|%$|^(verdadeiro|falso|true|false)$/i.test(v) ? "'" + v : v;
        else if (v?.richText) cel.e = v.richText.map((t) => t.text).join("");
        else if (v?.text !== undefined) cel.e = typeof v.text === "string" ? v.text : v.text?.richText?.map((t) => t.text).join("") ?? "";
        else if (v?.error) cel.e = erroDoExcel(v.error).codigo;
        else if (v?.result !== undefined) cel.e = typeof v.result === "number" ? paraTexto(v.result) : String(v.result);
        if (fmt) cel.fmt = fmt;
        if (x.font?.bold) cel.b = true;
        if (x.font?.italic) cel.i = true;
        if (x.font?.size && x.font.size !== 11) cel.tam = x.font.size;
        if (x.font?.name && !/^calibri$/i.test(x.font.name)) cel.fn = x.font.name;
        if (x.alignment?.wrapText) cel.qb = true;
        const bd = ["top", "bottom", "left", "right"].filter((l) => x.border?.[l]?.style).map((l) => l[0]).join("");
        if (bd) cel.bd = bd;
        if (hexDeArgb(x.font?.color) && x.font.color.argb !== "FF000000") cel.cor = hexDeArgb(x.font.color);
        if (["left", "center", "right"].includes(x.alignment?.horizontal)) cel.al = x.alignment.horizontal;
        if (x.fill?.type === "pattern" && x.fill.pattern === "solid" && hexDeArgb(x.fill.fgColor)) cel.bg = hexDeArgb(x.fill.fgColor);
        if (cel.e === undefined && Object.keys(cel).length === 0) return;
        p.cel[chave(rn - 1, cn - 1)] = cel;
        maxR = Math.max(maxR, rn);
        maxC = Math.max(maxC, cn);
      });
    });
    (ws.columns || []).forEach((col, i) => { if (col?.width) p.larg[i] = Math.round(col.width * 7); });
    for (const rg of Object.values(ws._merges || {})) {
      const m = rg.model;
      if (m) p.mescl.push({ r1: m.top - 1, c1: m.left - 1, r2: m.bottom - 1, c2: m.right - 1 });
    }
    const vista = ws.views?.[0];
    if (vista?.state === "frozen" && (vista.xSplit || vista.ySplit)) p.cong = { r: vista.ySplit || 0, c: vista.xSplit || 0 };
    const af = typeof ws.autoFilter === "string" ? ws.autoFilter : ws.autoFilter?.from && ws.autoFilter?.to ? `${ws.autoFilter.from}:${ws.autoFilter.to}` : null;
    const m = af && /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?\d+$/i.exec(af);
    if (m) p.filtro = { r: Number(m[2]) - 1, c1: indiceColuna(m[1].toUpperCase()), c2: indiceColuna(m[3].toUpperCase()), excl: {} };
    p.linhas = Math.max(100, maxR + 20);
    p.colunas = Math.max(26, maxC + 2);
    planilhas.push(p);
  });
  return planilhas.length ? planilhas : [novaAba()];
}

// ---------- CSV ----------

function lerCsv(texto) {
  texto = texto.replace(/^﻿/, "");
  const primeira = texto.split(/\r?\n/)[0];
  const sep = (primeira.match(/;/g) || []).length >= (primeira.match(/,/g) || []).length ? ";" : primeira.includes("\t") ? "\t" : ",";
  const linhas = [];
  let linha = [], campo = "", aspas = false;
  for (let i = 0; i < texto.length; i++) {
    const ch = texto[i];
    if (aspas) {
      if (ch === '"' && texto[i + 1] === '"') { campo += '"'; i++; }
      else if (ch === '"') aspas = false;
      else campo += ch;
    } else if (ch === '"') aspas = true;
    else if (ch === sep) { linha.push(campo); campo = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && texto[i + 1] === "\n") i++;
      linha.push(campo); linhas.push(linha); linha = []; campo = "";
    } else campo += ch;
  }
  if (campo || linha.length) { linha.push(campo); linhas.push(linha); }
  return linhas;
}

function csvDe(p, a, valor) {
  let maxR = -1, maxC = -1;
  for (const k of Object.keys(p.cel)) { const { r, c } = lerChave(k); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); }
  const esc = (s) => (/[;"\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const linhas = [];
  for (let r = 0; r <= maxR; r++) {
    const l = [];
    for (let c = 0; c <= maxC; c++) {
      const cel = p.cel[chave(r, c)];
      const v = valor(a, r, c);
      l.push(esc(cel?.fmt === "data" || ehFormulaData(cel?.e) ? exibir(v, "data") : v instanceof Erro ? v.codigo : paraTexto(v)));
    }
    linhas.push(l.join(";"));
  }
  return "﻿" + linhas.join("\r\n");
}

// ---------- Estilos ----------

const sBotao = (ativo) => ({ minWidth: 30, height: 28, padding: "0 7px", border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: ativo ? "rgba(154,110,30,0.2)" : "#fff", cursor: "pointer", fontSize: 13, fontFamily: "sans-serif", color: "#2a2a2a", whiteSpace: "nowrap" });
const sSelect = { height: 28, border: "1px solid rgba(0,0,0,0.12)", borderRadius: 6, background: "#fff", fontSize: 13, fontFamily: "sans-serif", padding: "0 4px" };
const sSep = { width: 1, height: 20, background: "rgba(0,0,0,0.15)", margin: "0 3px" };
const sCab = { position: "sticky", background: "#efece4", color: "#555", fontWeight: 500, fontSize: 12, textAlign: "center", border: "1px solid #d6d2c8", userSelect: "none", padding: 0 };

// ---------- Componente ----------

export default function Planilha() {
  const [pastas, setPastas] = useState(() => {
    const salvas = ler(CHAVE, null);
    return salvas?.length ? salvas : [novaPasta()];
  });
  const [atualId, setAtualId] = useState(() => pastas[0].id);
  const [aba, setAba] = useState(0);
  const [sel, setSel] = useState({ r: 0, c: 0 });
  const [anc, setAnc] = useState({ r: 0, c: 0 });
  const [editando, setEditando] = useState(null); // null | "celula" | "barra"
  const [rascunho, setRascunho] = useState("");
  const [msg, setMsg] = useState("");
  const [nomeCaixa, setNomeCaixa] = useState(null);
  const [menuFiltro, setMenuFiltro] = useState(null); // { c, x, y, marcados: Set, busca }
  const hist = useRef({ passado: [], futuro: [] });
  const clip = useRef(null);
  const grade = useRef(null);
  const arquivo = useRef(null);
  const arrastando = useRef(false);
  const modoDigitar = useRef(false);
  const emEdicao = useRef(false); // evita confirmar duas vezes (tecla + perda de foco)
  const [rolagem, setRolagem] = useState({ top: 0, altura: 800 });
  const [preench, setPreench] = useState(null); // área de destino da alça de preenchimento
  const preenchRef = useRef(null);
  const medidas = useRef(new Map()); // alturas reais das linhas já desenhadas
  const larguras = useRef(new Map());
  const cabAlt = useRef(26);
  const offsRef = useRef(null);
  const [, redesenhar] = useState(0);

  const pasta = pastas.find((x) => x.id === atualId) || pastas[0];
  const planilha = pasta.planilhas[Math.min(aba, pasta.planilhas.length - 1)];
  const abaIdx = pasta.planilhas.indexOf(planilha);
  const valor = useMemo(() => criarCalculadora(pasta.planilhas), [pasta.planilhas]);
  const ocultas = useMemo(() => linhasOcultas(planilha, valor, abaIdx), [planilha, valor, abaIdx]);
  const filtro = planilha.filtro;
  const mescl = planilha.mescl || [];
  const cong = planilha.cong || { r: 0, c: 0 };
  const { mestres, cobertas } = useMemo(() => {
    const mestres = new Map(), cobertas = new Map();
    for (const m of mescl) {
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) (r === m.r1 && c === m.c1 ? mestres : cobertas).set(chave(r, c), m);
    }
    return { mestres, cobertas };
  }, [mescl]);
  const mesclDe = (r, c) => mestres.get(chave(r, c)) || cobertas.get(chave(r, c));

  useEffect(() => {
    if (!gravar(CHAVE, pastas)) setMsg("Atenção: o arquivo é grande demais para ficar salvo neste navegador. Salve em Excel (.xlsx) para não perder.");
  }, [pastas]);

  useEffect(() => {
    const el = grade.current?.querySelector(`[data-k="${chave(sel.r, sel.c)}"]`);
    if (el) el.scrollIntoView({ block: "nearest", inline: "nearest" });
    else if (grade.current && offsRef.current) grade.current.scrollTop = Math.max(0, offsRef.current[sel.r] - grade.current.clientHeight / 2);
  }, [sel]);

  // Medidas reais (linhas com texto quebrado crescem): usadas para desenhar só o que aparece na tela.
  useLayoutEffect(() => {
    medidas.current = new Map();
    larguras.current = new Map();
  }, [pasta.id, abaIdx]);
  useLayoutEffect(() => {
    const g = grade.current;
    if (!g) return;
    let mudou = false;
    let base = Infinity;
    g.querySelectorAll("tr[data-r]").forEach((tr) => {
      const r = Number(tr.dataset.r);
      const h = tr.offsetHeight;
      if (!h) return;
      if (!planilha.alt?.[r]) base = Math.min(base, h);
      if (medidas.current.get(r) !== h) { medidas.current.set(r, h); mudou = true; }
    });
    if (base !== Infinity && medidas.current.get("base") !== base) { medidas.current.set("base", base); mudou = true; }
    g.querySelectorAll("th[data-col]").forEach((th) => {
      const c = Number(th.dataset.col);
      if (larguras.current.get(c) !== th.offsetWidth) { larguras.current.set(c, th.offsetWidth); mudou = true; }
    });
    const cab = g.querySelector("thead tr")?.offsetHeight;
    if (cab && cab !== cabAlt.current) { cabAlt.current = cab; mudou = true; }
    if (g.clientHeight && g.clientHeight !== rolagem.altura) { setRolagem((x) => ({ ...x, altura: g.clientHeight })); return; }
    if (mudou) redesenhar((x) => x + 1);
  });

  const r1 = Math.min(sel.r, anc.r), r2 = Math.max(sel.r, anc.r);
  const c1 = Math.min(sel.c, anc.c), c2 = Math.max(sel.c, anc.c);
  const dentro = (r, c) => r >= r1 && r <= r2 && c >= c1 && c <= c2;
  const celSel = planilha.cel[chave(sel.r, sel.c)] || {};
  const refSel = r1 === r2 && c1 === c2 ? chave(sel.r, sel.c) : `${chave(r1, c1)}:${chave(r2, c2)}`;

  // ----- alterações com histórico -----

  function mudarPlanilhas(fn) {
    setPastas((ps) => ps.map((x) => {
      if (x.id !== pasta.id) return x;
      const novas = fn(x.planilhas);
      if (novas === x.planilhas) return x;
      const h = hist.current.passado;
      if (h[h.length - 1] !== x.planilhas) h.push(x.planilhas);
      if (hist.current.passado.length > 100) hist.current.passado.shift();
      hist.current.futuro = [];
      return { ...x, planilhas: novas, atualizado: Date.now() };
    }));
  }

  function desfazer(refazer = false) {
    const de = refazer ? hist.current.futuro : hist.current.passado;
    const para = refazer ? hist.current.passado : hist.current.futuro;
    if (!de.length) return;
    const anterior = de.pop();
    setPastas((ps) => ps.map((x) => {
      if (x.id !== pasta.id) return x;
      if (para[para.length - 1] !== x.planilhas) para.push(x.planilhas);
      return { ...x, planilhas: anterior };
    }));
  }

  // edicoes: [{ r, c, patch }] — patch mescla na célula; célula vazia sem estilo é removida.
  function editar(edicoes, alvo = abaIdx) {
    mudarPlanilhas((pls) => pls.map((p, i) => {
      if (i !== alvo) return p;
      const cel = { ...p.cel };
      let linhas = p.linhas, colunas = p.colunas;
      for (const { r, c, patch } of edicoes) {
        if (r < 0 || c < 0) continue;
        const k = chave(r, c);
        const nova = { ...cel[k], ...patch };
        for (const [a, v] of Object.entries(nova)) if (v === undefined || v === false || v === "") delete nova[a];
        if (Object.keys(nova).length) cel[k] = nova; else delete cel[k];
        linhas = Math.max(linhas, r + 1);
        colunas = Math.max(colunas, c + 1);
      }
      return { ...p, cel, linhas, colunas };
    }));
  }

  function noIntervalo(fn) {
    const ed = [];
    for (let r = r1; r <= r2; r++) {
      if (ocultas.has(r)) continue; // como no Excel, linhas filtradas não são alteradas
      for (let c = c1; c <= c2; c++) ed.push({ r, c, patch: fn(planilha.cel[chave(r, c)] || {}) });
    }
    editar(ed);
  }

  // ----- edição de célula -----

  function comecarEdicao(texto, digitando = false, onde = "celula") {
    modoDigitar.current = digitando;
    emEdicao.current = true;
    // Monta o campo de edição na hora, para a próxima tecla digitada já cair nele.
    flushSync(() => {
      setRascunho(texto);
      setEditando(onde);
    });
  }

  function confirmar(dr = 0, dc = 0) {
    if (!editando || !emEdicao.current) return;
    emEdicao.current = false;
    const atualE = celSel.e || "";
    if (rascunho !== atualE) {
      const patch = { e: rascunho };
      if (rascunho.includes("\n")) patch.qb = true;
      if (!celSel.fmt && rascunho[0] !== "=" && rascunho[0] !== "'") {
        const { fmt } = interpretarEntrada(rascunho);
        if (fmt) patch.fmt = fmt;
      }
      editar([{ r: sel.r, c: sel.c, patch }]);
    }
    grade.current?.focus(); // antes de desmontar o campo, para não perder teclas
    setEditando(null);
    mover(dr, dc, false);
  }

  function cancelar() {
    emEdicao.current = false;
    grade.current?.focus();
    setEditando(null);
  }

  function mover(dr, dc, estender) {
    // Saindo de uma célula mesclada, anda a partir da borda dela.
    const m0 = !estender && mesclDe(sel.r, sel.c);
    const base = { r: sel.r, c: sel.c };
    if (m0) { base.r = dr > 0 ? m0.r2 : m0.r1; base.c = dc > 0 ? m0.c2 : m0.c1; }
    let r = Math.max(0, Math.min(planilha.linhas - 1, base.r + dr));
    if (dr && ocultas.has(r)) {
      const passo = Math.sign(dr);
      while (ocultas.has(r) && r + passo >= 0 && r + passo < planilha.linhas) r += passo;
      if (ocultas.has(r)) r = sel.r;
    }
    let c = Math.max(0, Math.min(planilha.colunas - 1, base.c + dc));
    const m = !estender && mesclDe(r, c);
    if (m) { r = m.r1; c = m.c1; }
    setSel({ r, c });
    if (!estender) setAnc({ r, c });
  }

  // Ao editar uma fórmula, clicar numa célula insere a referência dela.
  const aceitaRef = () => editando && rascunho[0] === "=" && /[=(;,+\-*/^&<>:]$/.test(rascunho.trimEnd());

  function mouseDownCelula(e, r, c) {
    if (aceitaRef()) {
      e.preventDefault();
      setRascunho((s) => s + chave(r, c));
      return;
    }
    if (editando) confirmar();
    arrastando.current = true;
    setSel({ r, c });
    if (!e.shiftKey) setAnc({ r, c });
    grade.current?.focus();
  }

  function teclaEdicao(e) {
    if (e.key === "Enter" && e.altKey) {
      // Alt+Enter: nova linha dentro da célula, como no Excel
      e.preventDefault();
      const t = e.target, a = t.selectionStart, b = t.selectionEnd;
      setRascunho((x) => x.slice(0, a) + "\n" + x.slice(b));
      requestAnimationFrame(() => t.setSelectionRange?.(a + 1, a + 1));
      return;
    }
    if (e.key === "Enter") { e.preventDefault(); confirmar(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === "Tab") { e.preventDefault(); confirmar(0, e.shiftKey ? -1 : 1); }
    else if (e.key === "Escape") { e.preventDefault(); cancelar(); }
    else if (modoDigitar.current && rascunho[0] !== "=" && e.key.startsWith("Arrow")) {
      e.preventDefault();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      confirmar(...d);
    }
  }

  function teclaGrade(e) {
    if (editando) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const setas = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (setas[k]) {
      e.preventDefault();
      let [dr, dc] = setas[k];
      if (ctrl) { dr *= 100000; dc *= 100000; }
      mover(dr, dc, e.shiftKey);
    } else if (k === "Enter") { e.preventDefault(); mover(e.shiftKey ? -1 : 1, 0); }
    else if (k === "Tab") { e.preventDefault(); mover(0, e.shiftKey ? -1 : 1); }
    else if (k === "Home") { e.preventDefault(); mover(0, -100000, e.shiftKey); }
    else if (k === "F2") { e.preventDefault(); comecarEdicao(celSel.e || ""); }
    else if (k === "Delete" || k === "Backspace") { e.preventDefault(); noIntervalo(() => ({ e: undefined })); }
    else if (ctrl && k.toLowerCase() === "z") { e.preventDefault(); desfazer(e.shiftKey); }
    else if (ctrl && k.toLowerCase() === "y") { e.preventDefault(); desfazer(true); }
    else if (ctrl && k.toLowerCase() === "b") { e.preventDefault(); alternar("b"); }
    else if (ctrl && k.toLowerCase() === "i") { e.preventDefault(); alternar("i"); }
    else if (ctrl && k.toLowerCase() === "d") { e.preventDefault(); preencher("baixo"); }
    else if (ctrl && k.toLowerCase() === "r") { e.preventDefault(); preencher("direita"); }
    else if (ctrl && k.toLowerCase() === "a") { e.preventDefault(); setAnc({ r: planilha.linhas - 1, c: planilha.colunas - 1 }); setSel({ r: 0, c: 0 }); }
    else if (!ctrl && !e.altKey && k.length === 1) { e.preventDefault(); comecarEdicao(k, true); }
  }

  // ----- formatação -----

  function alternar(prop) {
    const ligar = !celSel[prop];
    noIntervalo(() => ({ [prop]: ligar }));
  }
  const formatar = (patch) => noIntervalo(() => patch);

  function preencher(direcao) {
    const ed = [];
    if (direcao === "baixo") {
      for (let c = c1; c <= c2; c++) {
        const origem = planilha.cel[chave(r1, c)] || {};
        for (let r = r1 + 1; r <= r2; r++) ed.push({ r, c, patch: copiaDeslocada(origem, r - r1, 0) });
      }
    } else {
      for (let r = r1; r <= r2; r++) {
        const origem = planilha.cel[chave(r, c1)] || {};
        for (let c = c1 + 1; c <= c2; c++) ed.push({ r, c, patch: copiaDeslocada(origem, 0, c - c1) });
      }
    }
    if (ed.length) editar(ed);
  }

  function copiaDeslocada(cel, dr, dc) {
    const nova = { ...CEL_VAZIA, ...cel };
    if (cel.e?.[0] === "=") nova.e = "=" + deslocarFormula(cel.e.slice(1), dr, dc);
    return nova;
  }

  function autoSoma() {
    const numero = (r, c) => typeof valor(abaIdx, r, c) === "number";
    let r = sel.r - 1;
    while (r >= 0 && numero(r, sel.c)) r--;
    let f;
    if (r < sel.r - 1) f = `=SOMA(${chave(r + 1, sel.c)}:${chave(sel.r - 1, sel.c)})`;
    else {
      let c = sel.c - 1;
      while (c >= 0 && numero(sel.r, c)) c--;
      f = c < sel.c - 1 ? `=SOMA(${chave(sel.r, c + 1)}:${chave(sel.r, sel.c - 1)})` : "=SOMA()";
    }
    comecarEdicao(f, false, "celula");
  }

  // ----- linhas, colunas, abas -----

  function estrutura(eixo, delta) {
    const pos = eixo === "r" ? r1 : c1;
    const qtd = eixo === "r" ? r2 - r1 + 1 : c2 - c1 + 1;
    mudarPlanilhas((pls) => {
      let x = pls;
      for (let n = 0; n < qtd; n++) x = reestruturar(x, abaIdx, eixo, pos, delta);
      return x;
    });
    if (delta < 0) { setAnc(sel); }
  }

  function adicionarAba() {
    let n = pasta.planilhas.length + 1;
    while (pasta.planilhas.some((p) => p.nome === `Planilha${n}`)) n++;
    mudarPlanilhas((pls) => [...pls, novaAba(`Planilha${n}`)]);
    setAba(pasta.planilhas.length);
  }

  function renomearAba(i) {
    const antigo = pasta.planilhas[i].nome;
    const novo = prompt("Nome da planilha:", antigo)?.trim();
    if (!novo || novo === antigo) return;
    if (pasta.planilhas.some((p, j) => j !== i && p.nome.toUpperCase() === novo.toUpperCase())) return setMsg("Já existe uma planilha com esse nome.");
    mudarPlanilhas((pls) => pls.map((p, j) => {
      const cel = {};
      for (const [k, v] of Object.entries(p.cel)) {
        cel[k] = v.e?.[0] === "=" ? { ...v, e: "=" + transformarRefs(v.e.slice(1), (t) => { if (t.planilha?.toUpperCase() === antigo.toUpperCase()) t.planilha = novo; }) } : v;
      }
      return { ...p, cel, nome: j === i ? novo : p.nome };
    }));
  }

  function excluirAba() {
    if (pasta.planilhas.length < 2 || !confirm(`Excluir a planilha "${planilha.nome}"?`)) return;
    mudarPlanilhas((pls) => pls.filter((_, j) => j !== abaIdx));
    setAba(Math.max(0, abaIdx - 1));
  }

  // ----- pastas (arquivos) -----

  function trocarPasta(id) {
    hist.current = { passado: [], futuro: [] };
    setAtualId(id);
    setAba(0);
    setSel({ r: 0, c: 0 });
    setAnc({ r: 0, c: 0 });
    setEditando(null);
  }

  function criarPasta(nome, planilhas) {
    const n = novaPasta(nome, planilhas);
    setPastas((ps) => [n, ...ps]);
    trocarPasta(n.id);
  }

  function excluirPasta() {
    if (!confirm(`Excluir "${pasta.nome}"?`)) return;
    const resto = pastas.filter((x) => x.id !== pasta.id);
    const lista = resto.length ? resto : [novaPasta()];
    setPastas(lista);
    trocarPasta(lista[0].id);
  }

  async function abrirArquivo(e) {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    const nome = f.name.replace(/\.[^.]+$/, "");
    try {
      setMsg("Abrindo arquivo...");
      if (/\.xlsx$/i.test(f.name)) criarPasta(nome, await importarXlsx(f));
      else if (/\.(csv|txt|tsv)$/i.test(f.name)) {
        const p = novaAba(nome.slice(0, 31));
        const linhas = lerCsv(await f.text());
        linhas.forEach((l, r) => l.forEach((v, c) => {
          if (v === "") return;
          const { fmt } = interpretarEntrada(v);
          p.cel[chave(r, c)] = fmt ? { e: v, fmt } : { e: v };
        }));
        p.linhas = Math.max(100, linhas.length + 20);
        p.colunas = Math.max(26, Math.max(0, ...linhas.map((l) => l.length)) + 2);
        criarPasta(nome, [p]);
      } else throw new Error("Formato não suportado. Use .xlsx ou .csv (arquivos .xls antigos: salve como .xlsx no Excel antes).");
      setMsg("");
    } catch (err) {
      setMsg(err.message || "Não foi possível abrir o arquivo.");
    }
  }

  async function salvarXlsx() {
    setMsg("Gerando .xlsx...");
    try { await exportarXlsx(pasta, valor); setMsg(""); } catch (err) { setMsg("Erro ao gerar .xlsx: " + err.message); }
  }

  function salvarCsv() {
    baixar(new Blob([csvDe(planilha, abaIdx, valor)], { type: "text/csv;charset=utf-8" }), `${pasta.nome} - ${planilha.nome}.csv`);
  }

  // ----- ordenar e filtrar -----

  // Ordena as linhas r1..r2 (colunas c1..c2) pela coluna col. Fórmulas são ajustadas
  // como se a linha tivesse sido movida, igual ao Excel.
  function ordenarLinhas(a, col, desc) {
    const linhasOrd = [];
    for (let r = a.r1; r <= a.r2; r++) linhasOrd.push({ r, v: valor(abaIdx, r, col) });
    linhasOrd.sort((x, y) => compararOrdem(x.v, y.v, desc) || x.r - y.r);
    const ed = [];
    linhasOrd.forEach(({ r: origem }, i) => {
      const destino = a.r1 + i;
      if (origem === destino) return;
      for (let c = a.c1; c <= a.c2; c++) ed.push({ r: destino, c, patch: copiaDeslocada(planilha.cel[chave(origem, c)] || {}, destino - origem, 0) });
    });
    if (ed.length) editar(ed);
  }

  function ordenar(desc) {
    if (editando) confirmar();
    const col = sel.c;
    // Dentro do filtro: ordena os dados abaixo dos títulos.
    if (filtro && col >= filtro.c1 && col <= filtro.c2 && sel.r >= filtro.r) {
      ordenarLinhas({ r1: filtro.r + 1, r2: fimFiltro(planilha, filtro), c1: filtro.c1, c2: filtro.c2 }, col, desc);
      setMsg("");
      return;
    }
    const a = r1 === r2 && c1 === c2 ? regiaoAtual(planilha, sel.r, sel.c) : { r1, r2, c1, c2 };
    if (a.r1 === a.r2) return setMsg("Não há dados para ordenar: selecione uma célula dentro da tabela.");
    // Primeira linha é título se for toda texto e (estiver em negrito ou houver números/datas abaixo dela).
    const primeira = [];
    for (let c = a.c1; c <= a.c2; c++) primeira.push(valor(abaIdx, a.r1, c));
    let titulo = primeira.every((v) => v === null || typeof v === "string") && primeira.some((v) => typeof v === "string");
    if (titulo) {
      let negrito = false, numeroAbaixo = false;
      for (let c = a.c1; c <= a.c2; c++) if (planilha.cel[chave(a.r1, c)]?.b) negrito = true;
      for (let r = a.r1 + 1; r <= a.r2 && !numeroAbaixo; r++) for (let c = a.c1; c <= a.c2; c++) if (typeof valor(abaIdx, r, c) === "number") numeroAbaixo = true;
      titulo = negrito || numeroAbaixo;
    }
    ordenarLinhas({ ...a, r1: titulo ? a.r1 + 1 : a.r1 }, col, desc);
    setMsg(`Ordenado por ${nomeColuna(col)} (${desc ? "Z→A" : "A→Z"}), linhas ${(titulo ? a.r1 + 1 : a.r1) + 1} a ${a.r2 + 1}${titulo ? "; a linha " + (a.r1 + 1) + " ficou como título" : ""}. Ctrl+Z desfaz.`);
  }

  function alternarFiltro() {
    if (editando) confirmar();
    setMenuFiltro(null);
    if (filtro) {
      mudarPlanilhas((pls) => pls.map((p, i) => (i === abaIdx ? { ...p, filtro: undefined } : p)));
      return;
    }
    const a = r1 === r2 && c1 === c2 ? regiaoAtual(planilha, sel.r, sel.c) : { r1, r2, c1, c2 };
    if (!planilha.cel[chave(a.r1, a.c1)]?.e && a.r1 === a.r2 && a.c1 === a.c2) return setMsg("Selecione uma célula dentro da tabela para filtrar.");
    mudarPlanilhas((pls) => pls.map((p, i) => (i === abaIdx ? { ...p, filtro: { r: a.r1, c1: a.c1, c2: a.c2, excl: {} } } : p)));
    setMsg("");
  }

  function abrirMenuFiltro(e, c) {
    e.preventDefault();
    e.stopPropagation();
    if (editando) confirmar();
    const caixa = e.currentTarget.getBoundingClientRect();
    const excl = new Set(filtro.excl?.[c] || []);
    const fim = fimFiltro(planilha, filtro);
    const vistos = new Map();
    for (let r = filtro.r + 1; r <= fim; r++) {
      const t = textoFiltro(planilha, valor, abaIdx, r, c);
      if (!vistos.has(t)) vistos.set(t, valor(abaIdx, r, c));
    }
    for (const t of excl) if (!vistos.has(t)) vistos.set(t, t); // mantém valores já desmarcados
    const itens = [...vistos.entries()].sort((x, y) => compararOrdem(x[1], y[1], false)).map(([t]) => t);
    setMenuFiltro({
      c, itens, busca: "",
      marcados: new Set(itens.filter((t) => !excl.has(t))),
      x: Math.min(caixa.left, window.innerWidth - 270), y: caixa.bottom + 2,
    });
  }

  function aplicarFiltro(excluidos) {
    const { c } = menuFiltro;
    mudarPlanilhas((pls) => pls.map((p, i) => {
      if (i !== abaIdx || !p.filtro) return p;
      const excl = { ...p.filtro.excl };
      if (excluidos.length) excl[c] = excluidos; else delete excl[c];
      return { ...p, filtro: { ...p.filtro, excl } };
    }));
    setMenuFiltro(null);
  }

  // ----- área de transferência -----

  function copiar(e, recortar = false) {
    if (editando) return;
    e.preventDefault();
    const linhas = [];
    const brutos = [];
    for (let r = r1; r <= r2; r++) {
      if (ocultas.has(r)) continue;
      const l = [], b = [];
      for (let c = c1; c <= c2; c++) {
        const cel = planilha.cel[chave(r, c)] || {};
        b.push(cel);
        const v = valor(abaIdx, r, c);
        l.push(cel.e?.[0] === "=" ? exibir(v, cel.fmt) : cel.e || "");
      }
      linhas.push(l.join("\t"));
      brutos.push(b);
    }
    const texto = linhas.join("\n");
    e.clipboardData.setData("text/plain", texto);
    clip.current = { texto, brutos, r: r1, c: c1, aba: abaIdx, recortar };
  }

  function colar(e) {
    if (editando) return;
    e.preventDefault();
    const texto = e.clipboardData.getData("text/plain");
    const ed = [];
    const interno = clip.current && clip.current.texto === texto ? clip.current : null;
    if (interno) {
      interno.brutos.forEach((l, i) => l.forEach((cel, j) => {
        ed.push({ r: sel.r + i, c: sel.c + j, patch: interno.recortar ? { ...CEL_VAZIA, ...cel } : copiaDeslocada(cel, sel.r - interno.r, sel.c - interno.c) });
      }));
      if (interno.recortar) {
        const limpar = [];
        interno.brutos.forEach((l, i) => l.forEach((_, j) => {
          const r = interno.r + i, c = interno.c + j;
          if (!(interno.aba === abaIdx && r >= sel.r && r < sel.r + interno.brutos.length && c >= sel.c && c < sel.c + l.length)) {
            limpar.push({ r, c, patch: { ...CEL_VAZIA } });
          }
        }));
        if (limpar.length) editar(limpar, interno.aba);
        clip.current = null;
      }
    } else {
      texto.replace(/\r?\n$/, "").split(/\r?\n/).forEach((l, i) => l.split("\t").forEach((v, j) => {
        const { fmt } = interpretarEntrada(v);
        ed.push({ r: sel.r + i, c: sel.c + j, patch: fmt ? { e: v, fmt } : { e: v } });
      }));
    }
    editar(ed);
    const fim = ed[ed.length - 1];
    if (fim) setAnc({ r: fim.r, c: fim.c });
  }

  // ----- largura das colunas -----

  function redimensionar(e, c) {
    e.preventDefault();
    e.stopPropagation();
    const x0 = e.clientX;
    const w0 = planilha.larg?.[c] || LARG_PADRAO;
    let w = w0;
    const mov = (ev) => {
      w = Math.max(30, w0 + ev.clientX - x0);
      const th = grade.current?.querySelector(`[data-col="${c}"]`);
      if (th) th.style.width = th.style.minWidth = th.style.maxWidth = w + "px";
    };
    const fim = () => {
      window.removeEventListener("mousemove", mov);
      window.removeEventListener("mouseup", fim);
      if (w !== w0) mudarPlanilhas((pls) => pls.map((p, i) => (i === abaIdx ? { ...p, larg: { ...p.larg, [c]: w } } : p)));
    };
    window.addEventListener("mousemove", mov);
    window.addEventListener("mouseup", fim);
  }

  useEffect(() => {
    const solta = () => { arrastando.current = false; };
    window.addEventListener("mouseup", solta);
    return () => window.removeEventListener("mouseup", solta);
  }, []);

  // ----- mesclar, bordas, congelar -----

  const mudarAba = (fn) => mudarPlanilhas((pls) => pls.map((p, i) => (i === abaIdx ? fn(p) : p)));

  function mesclar() {
    if (editando) confirmar();
    const sobre = mescl.filter((m) => !(m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2));
    if (sobre.length) return mudarAba((p) => ({ ...p, mescl: (p.mescl || []).filter((m) => !sobre.includes(m)) }));
    if (r1 === r2 && c1 === c2) return setMsg("Selecione duas ou mais células para mesclar.");
    const outras = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if ((r !== r1 || c !== c1) && planilha.cel[chave(r, c)]?.e) outras.push(chave(r, c));
    if (outras.length && !confirm("Mesclar mantém só o conteúdo da célula superior esquerda. Continuar?")) return;
    mudarAba((p) => {
      const cel = { ...p.cel };
      for (const k of outras) { const { e, ...resto } = cel[k]; if (Object.keys(resto).length) cel[k] = resto; else delete cel[k]; }
      const k0 = chave(r1, c1);
      cel[k0] = { ...cel[k0], al: cel[k0]?.al || "center" };
      return { ...p, cel, mescl: [...(p.mescl || []), { r1, c1, r2, c2 }] };
    });
    setSel({ r: r1, c: c1 });
    setAnc({ r: r1, c: c1 });
  }

  function bordas(tipo) {
    const ed = [];
    for (let r = r1; r <= r2; r++) {
      if (ocultas.has(r)) continue;
      for (let c = c1; c <= c2; c++) {
        const atual = planilha.cel[chave(r, c)]?.bd || "";
        // Entre duas células a linha é desenhada só uma vez (embaixo/à direita).
        const lados = {
          todas: "br" + (r === r1 ? "t" : "") + (c === c1 ? "l" : ""),
          externas: (r === r1 ? "t" : "") + (r === r2 ? "b" : "") + (c === c1 ? "l" : "") + (c === c2 ? "r" : ""),
          inferior: r === r2 ? "b" : "", superior: r === r1 ? "t" : "", esquerda: c === c1 ? "l" : "", direita: c === c2 ? "r" : "",
        }[tipo] || "";
        const novo = tipo === "nenhuma" ? "" : [..."tblr"].filter((l) => atual.includes(l) || lados.includes(l)).join("");
        if (novo !== atual) ed.push({ r, c, patch: { bd: novo || undefined } });
      }
    }
    if (ed.length) editar(ed);
  }

  function congelar(tipo) {
    const novo = { paineis: { r: sel.r, c: sel.c }, linha: { r: 1, c: 0 }, coluna: { r: 0, c: 1 } }[tipo];
    mudarAba((p) => ({ ...p, cong: novo && (novo.r || novo.c) ? novo : undefined }));
  }

  function redimensionarLinha(e, r) {
    e.preventDefault();
    e.stopPropagation();
    const y0 = e.clientY;
    const tr = grade.current?.querySelector(`tr[data-r="${r}"]`);
    const h0 = tr?.offsetHeight || ALT_LINHA;
    let h = h0;
    const mov = (ev) => {
      h = Math.max(12, h0 + ev.clientY - y0);
      tr?.querySelectorAll("td,th").forEach((x) => { x.style.height = h + "px"; });
    };
    const fim = () => {
      window.removeEventListener("mousemove", mov);
      window.removeEventListener("mouseup", fim);
      if (h !== h0) mudarAba((p) => ({ ...p, alt: { ...p.alt, [r]: h } }));
    };
    window.addEventListener("mousemove", mov);
    window.addEventListener("mouseup", fim);
  }

  // ----- alça de preenchimento -----

  function aplicarPreenchimento(o, a) {
    if (!a?.dir) return;
    const vertical = a.dir === "baixo" || a.dir === "cima";
    const volta = a.dir === "cima" || a.dir === "esquerda";
    const qtd = vertical ? (a.r2 - a.r1) - (o.r2 - o.r1) : (a.c2 - a.c1) - (o.c2 - o.c1);
    const ed = [];
    const fixos = [];
    for (let x = vertical ? o.c1 : o.r1; x <= (vertical ? o.c2 : o.r2); x++) fixos.push(x);
    for (const fx of fixos) {
      const pos = [];
      for (let x = vertical ? o.r1 : o.c1; x <= (vertical ? o.r2 : o.c2); x++) pos.push(x);
      if (volta) pos.reverse();
      const origem = pos.map((x) => (vertical ? { r: x, c: fx } : { r: fx, c: x }));
      const cels = origem.map((x) => planilha.cel[chave(x.r, x.c)] || {});
      const serie = gerarSerie(cels, qtd);
      const passo = volta ? -1 : 1;
      const ultimo = pos[pos.length - 1];
      for (let t = 1; t <= qtd; t++) {
        const p = ultimo + passo * t;
        const alvo = vertical ? { r: p, c: fx } : { r: fx, c: p };
        const i = (t - 1) % cels.length;
        const patch = copiaDeslocada(cels[i], alvo.r - origem[i].r, alvo.c - origem[i].c);
        if (serie) patch.e = serie[t - 1];
        ed.push({ ...alvo, patch });
      }
    }
    if (ed.length) editar(ed);
    setAnc({ r: a.r1, c: a.c1 });
    setSel({ r: a.r2, c: a.c2 });
  }

  function iniciarPreenchimento(e) {
    e.preventDefault();
    e.stopPropagation();
    if (editando) confirmar();
    const o = { r1, r2, c1, c2 };
    preenchRef.current = { o, a: { ...o, dir: null } };
    setPreench(preenchRef.current.a);
    const fim = () => {
      window.removeEventListener("mouseup", fim);
      const atual = preenchRef.current;
      preenchRef.current = null;
      setPreench(null);
      if (atual) aplicarRef.current(atual.o, atual.a);
    };
    window.addEventListener("mouseup", fim);
  }
  const aplicarRef = useRef(null);
  aplicarRef.current = aplicarPreenchimento;

  function arrastarPreenchimento(r, c) {
    const { o } = preenchRef.current;
    const d = { baixo: r - o.r2, cima: o.r1 - r, direita: c - o.c2, esquerda: o.c1 - c };
    const dir = Object.keys(d).reduce((x, y) => (d[y] > d[x] ? y : x));
    const a = { ...o, dir: d[dir] > 0 ? dir : null };
    if (a.dir === "baixo") a.r2 = r; else if (a.dir === "cima") a.r1 = r; else if (a.dir === "direita") a.c2 = c; else if (a.dir === "esquerda") a.c1 = c;
    preenchRef.current.a = a;
    setPreench(a);
  }

  // Clique duplo na alça: preenche para baixo até onde a coluna vizinha tem dados.
  function preencherAteOFim(e) {
    e.preventDefault();
    e.stopPropagation();
    const tem = (r, c) => c >= 0 && !!planilha.cel[chave(r, c)]?.e;
    const viz = tem(r2 + 1, c1 - 1) ? c1 - 1 : tem(r2 + 1, c2 + 1) ? c2 + 1 : null;
    if (viz === null) return;
    let fim = r2;
    while (fim + 1 < planilha.linhas && tem(fim + 1, viz)) fim++;
    aplicarPreenchimento({ r1, r2, c1, c2 }, { r1, r2: fim, c1, c2, dir: "baixo" });
  }

  // ----- caixa de nome (ir para célula) -----

  function irPara(texto) {
    const m = /^\s*([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?\s*$/.exec(texto);
    setNomeCaixa(null);
    if (!m) return;
    const a = { r: Number(m[2]) - 1, c: indiceColuna(m[1]) };
    const b = m[3] ? { r: Number(m[4]) - 1, c: indiceColuna(m[3]) } : a;
    mudarPlanilhas((pls) => pls.map((p, i) => (i === abaIdx ? { ...p, linhas: Math.max(p.linhas, a.r + 1, b.r + 1), colunas: Math.max(p.colunas, a.c + 1, b.c + 1) } : p)));
    setAnc(a);
    setSel(b);
    grade.current?.focus();
  }

  // ----- resumo da seleção -----

  const resumo = useMemo(() => {
    if (r1 === r2 && c1 === c2) return null;
    let soma = 0, n = 0, cont = 0;
    for (let r = r1; r <= Math.min(r2, planilha.linhas - 1); r++) {
      for (let c = c1; c <= Math.min(c2, planilha.colunas - 1); c++) {
        const v = valor(abaIdx, r, c);
        if (v !== null && v !== "") cont++;
        if (typeof v === "number") { soma += v; n++; }
      }
    }
    return { soma, n, cont };
  }, [r1, r2, c1, c2, valor, abaIdx, planilha]);

  // ----- render -----

  const colunas = Array.from({ length: planilha.colunas }, (_, c) => c);
  const textoBarra = editando ? rascunho : celSel.e || "";

  // Desenha só as linhas visíveis (mais uma margem): planilhas grandes continuam leves.
  const base = medidas.current.get("base") || 25;
  const offs = new Array(planilha.linhas + 1);
  offs[0] = 0;
  for (let r = 0; r < planilha.linhas; r++) offs[r + 1] = offs[r] + (ocultas.has(r) ? 0 : medidas.current.get(r) ?? (planilha.alt?.[r] ? planilha.alt[r] + 1 : base));
  offsRef.current = offs;
  const linhaEm = (y) => { let a = 0, b = planilha.linhas - 1; while (a < b) { const m = (a + b + 1) >> 1; if (offs[m] <= y) a = m; else b = m - 1; } return a; };
  const nCong = Math.min(cong.r, planilha.linhas);
  let vIni = Math.max(nCong, linhaEm(rolagem.top - 600));
  const vFim = Math.min(planilha.linhas - 1, linhaEm(rolagem.top + rolagem.altura + 600));
  for (const m of mescl) if (m.r1 >= nCong && m.r1 < vIni && m.r2 >= vIni) vIni = m.r1;
  const linhasDesenho = [];
  for (let r = 0; r < nCong; r++) linhasDesenho.push(r);
  for (let r = vIni; r <= vFim; r++) linhasDesenho.push(r);
  const espacoAntes = offs[vIni] - offs[nCong];
  const espacoDepois = offs[planilha.linhas] - offs[vFim + 1];

  // Posições fixas dos painéis congelados.
  const topoCong = (r) => cabAlt.current + offs[r];
  const larguraCol = (c) => larguras.current.get(c) || (planilha.larg?.[c] || LARG_PADRAO) + 2;
  const esqCong = [];
  for (let c = 0, x = larguras.current.get("cab") || 46; c < cong.c; c++) { esqCong.push(x); x += larguraCol(c); }

  const noPreench = (r, c) => preench && r >= preench.r1 && r <= preench.r2 && c >= preench.c1 && c <= preench.c2 && !dentro(r, c);
  const mFim = mesclDe(r2, c2);
  const alca = mFim ? { r: mFim.r1, c: mFim.c1 } : { r: r2, c: c2 };

  function rolar(e) {
    const t = e.currentTarget.scrollTop;
    if (Math.abs(t - rolagem.top) > 200) setRolagem((x) => ({ ...x, top: t }));
  }

  const selecionarColuna = (e, c) => { e.preventDefault(); if (editando) confirmar(); setAnc({ r: planilha.linhas - 1, c: e.shiftKey ? anc.c : c }); setSel({ r: 0, c }); grade.current?.focus(); };
  const selecionarLinha = (e, r) => { e.preventDefault(); if (editando) confirmar(); setAnc({ r: e.shiftKey ? anc.r : r, c: planilha.colunas - 1 }); setSel({ r, c: 0 }); grade.current?.focus(); };

  function celula(r, c) {
    const k = chave(r, c);
    if (cobertas.has(k)) return null;
    const m = mestres.get(k);
    const span = m ? { rowSpan: m.r2 - m.r1 + 1, colSpan: m.c2 - m.c1 + 1 } : {};
    const cel = planilha.cel[k];
    const ativa = r === sel.r && c === sel.c;
    const congR = r < cong.r, congC = c < cong.c;
    const fixo = congR || congC ? { position: "sticky", top: congR ? topoCong(r) : undefined, left: congC ? esqCong[c] : undefined, zIndex: congR && congC ? 2 : 1 } : { position: "relative" };
    const altura = planilha.alt?.[r] || ALT_LINHA;
    if (ativa && editando === "celula") {
      return (
        <td key={c} data-k={k} {...span} style={{ ...fixo, zIndex: 4, padding: 0, border: "2px solid " + GOLD, height: altura }}>
          <textarea autoFocus value={rascunho} onChange={(e) => setRascunho(e.target.value)} onKeyDown={teclaEdicao}
            rows={Math.max(1, rascunho.split("\n").length)}
            onFocus={(e) => { const n = e.target.value.length; e.target.setSelectionRange(n, n); }}
            onBlur={() => { if (!aceitaRef()) confirmar(); }}
            style={{ position: "absolute", top: 0, left: 0, width: "100%", minWidth: 120, minHeight: "100%", border: 0, padding: "3px 4px", fontSize: cel?.tam ? cel.tam + "pt" : 13, fontFamily: cel?.fn || "inherit", lineHeight: 1.3, outline: "none", boxSizing: "border-box", zIndex: 1, background: "#fff", boxShadow: "0 2px 6px rgba(0,0,0,0.2)", resize: "none", overflow: "hidden" }} />
        </td>
      );
    }
    const v = cel ? valor(abaIdx, r, c) : null;
    const fmt = cel?.fmt || (ehFormulaData(cel?.e) ? "data" : undefined);
    const erro = v instanceof Erro;
    const al = cel?.al || (erro || typeof v === "boolean" ? "center" : typeof v === "number" ? "right" : "left");
    const sombras = [];
    const bd = cel?.bd || "";
    if (bd.includes("t")) sombras.push("inset 0 1px 0 #222");
    if (bd.includes("b")) sombras.push("inset 0 -1px 0 #222");
    if (bd.includes("l")) sombras.push("inset 1px 0 0 #222");
    if (bd.includes("r")) sombras.push("inset -1px 0 0 #222");
    if (cong.r && r + (m ? m.r2 - m.r1 : 0) === cong.r - 1) sombras.push("inset 0 -2px 0 #9a9a9a");
    if (cong.c && c + (m ? m.c2 - m.c1 : 0) === cong.c - 1) sombras.push("inset -2px 0 0 #9a9a9a");
    const marcada = dentro(r, c) && !ativa;
    const texto = cel ? exibir(v, fmt) : "";
    const cabFiltro = filtro && r === filtro.r && c >= filtro.c1 && c <= filtro.c2;
    return (
      <td key={c} data-k={k} {...span}
        onMouseDown={(e) => mouseDownCelula(e, r, c)}
        onMouseEnter={() => { if (preenchRef.current) arrastarPreenchimento(r, c); else if (arrastando.current) setSel({ r, c }); }}
        onDoubleClick={() => comecarEdicao(cel?.e || "")}
        title={erro ? "Erro na fórmula: " + v.codigo : undefined}
        style={{
          ...fixo,
          height: altura, padding: "0 4px", whiteSpace: cel?.qb ? "pre-wrap" : "pre", wordBreak: cel?.qb ? "break-word" : undefined,
          overflow: "hidden", textOverflow: cel?.qb ? undefined : "ellipsis", verticalAlign: cel?.qb ? "top" : "middle", lineHeight: cel?.qb ? 1.3 : undefined,
          borderRight: "1px solid #e4e1da", borderBottom: "1px solid #e4e1da",
          outline: ativa ? "2px solid " + GOLD : noPreench(r, c) ? "1px dashed #666" : "none", outlineOffset: ativa ? -2 : -1,
          backgroundColor: cel?.bg || "#fff",
          backgroundImage: marcada ? "linear-gradient(rgba(154,110,30,0.14), rgba(154,110,30,0.14))" : noPreench(r, c) ? "linear-gradient(rgba(0,0,0,0.06), rgba(0,0,0,0.06))" : undefined,
          boxShadow: sombras.join(", ") || undefined,
          color: erro ? "#c0392b" : cel?.cor || "#111", fontWeight: cel?.b ? "bold" : undefined, fontStyle: cel?.i ? "italic" : undefined,
          fontSize: cel?.tam ? cel.tam + "pt" : undefined, fontFamily: cel?.fn || undefined,
          textAlign: al, cursor: "cell", userSelect: "none",
        }}>
        {cabFiltro ? (
          <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
            <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", textAlign: al }}>{texto}</span>
            <button onMouseDown={(e) => abrirMenuFiltro(e, c)} title="Filtrar e ordenar esta coluna"
              style={{ flex: "none", width: 18, height: 18, padding: 0, fontSize: 10, lineHeight: "16px", border: "1px solid #bbb", borderRadius: 3, cursor: "pointer", background: filtro.excl?.[c]?.length ? GOLD : "#fff", color: filtro.excl?.[c]?.length ? "#fff" : "#444" }}>
              {filtro.excl?.[c]?.length ? "⏷" : "▾"}
            </button>
          </span>
        ) : texto}
        {r === alca.r && c === alca.c && !editando && (
          <div onMouseDown={iniciarPreenchimento} onDoubleClick={preencherAteOFim} title="Arraste para preencher (clique duplo: até o fim da tabela)"
            style={{ position: "absolute", right: 0, bottom: 0, width: 7, height: 7, background: GOLD, border: "1px solid #fff", cursor: "crosshair", zIndex: 3 }} />
        )}
      </td>
    );
  }

  return (
    <div style={{ fontFamily: "sans-serif" }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 10 }}>
        <select value={pasta.id} onChange={(e) => trocarPasta(e.target.value)} style={{ ...sSelect, maxWidth: 220 }} title="Planilhas salvas neste navegador">
          {pastas.map((x) => <option key={x.id} value={x.id}>{x.nome || "(sem nome)"}</option>)}
        </select>
        <input value={pasta.nome} onChange={(e) => setPastas((ps) => ps.map((x) => (x.id === pasta.id ? { ...x, nome: e.target.value } : x)))} style={{ ...sSelect, flex: "1 1 160px", padding: "0 8px" }} placeholder="Nome do arquivo" />
        <button style={sBotao()} onClick={() => criarPasta()}>＋ Nova</button>
        <button style={sBotao()} onClick={() => arquivo.current.click()}>📂 Abrir</button>
        <button style={sBotao()} onClick={salvarXlsx}>💾 Excel (.xlsx)</button>
        <button style={sBotao()} onClick={salvarCsv}>CSV</button>
        <button style={sBotao()} onClick={excluirPasta} title="Excluir arquivo">🗑️</button>
        <input ref={arquivo} type="file" accept=".xlsx,.csv,.tsv,.txt" onChange={abrirArquivo} style={{ display: "none" }} />
      </div>
      {msg && <p style={{ fontSize: 13, color: GOLD, margin: "0 0 8px" }}>{msg}</p>}

      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center", padding: 6, background: "#f4f1ea", border: "1px solid rgba(0,0,0,0.12)", borderRadius: "10px 10px 0 0" }}
        onMouseDown={(e) => { if (e.target.tagName === "BUTTON") e.preventDefault(); }}>
        <button style={sBotao()} onClick={() => desfazer()} title="Desfazer (Ctrl+Z)">↶</button>
        <button style={sBotao()} onClick={() => desfazer(true)} title="Refazer (Ctrl+Y)">↷</button>
        <span style={sSep} />
        <button style={{ ...sBotao(celSel.b), fontWeight: "bold" }} onClick={() => alternar("b")} title="Negrito (Ctrl+B)">N</button>
        <button style={{ ...sBotao(celSel.i), fontStyle: "italic" }} onClick={() => alternar("i")} title="Itálico (Ctrl+I)">I</button>
        <select style={{ ...sSelect, width: 110 }} value={celSel.fn || ""} onChange={(e) => formatar({ fn: e.target.value || undefined })} title="Fonte">
          <option value="">Fonte padrão</option>
          {["Calibri", "Arial", "Times New Roman", "Georgia", "Verdana", "Courier New"].map((f) => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
        </select>
        <select style={{ ...sSelect, width: 58 }} value={celSel.tam || ""} onChange={(e) => formatar({ tam: Number(e.target.value) || undefined })} title="Tamanho da fonte">
          <option value="">10</option>
          {[8, 9, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
        <label style={{ ...sBotao(), display: "inline-flex", alignItems: "center", gap: 2 }} title="Cor do texto">
          A<input type="color" value={celSel.cor || "#000000"} onChange={(e) => formatar({ cor: e.target.value })} style={{ width: 18, height: 18, border: 0, padding: 0, background: "none" }} />
        </label>
        <label style={{ ...sBotao(), display: "inline-flex", alignItems: "center", gap: 2 }} title="Cor de preenchimento">
          🪣<input type="color" value={celSel.bg || "#ffffff"} onChange={(e) => formatar({ bg: e.target.value === "#ffffff" ? undefined : e.target.value })} style={{ width: 18, height: 18, border: 0, padding: 0, background: "none" }} />
        </label>
        <span style={sSep} />
        <button style={sBotao(celSel.al === "left")} onClick={() => formatar({ al: celSel.al === "left" ? undefined : "left" })} title="Alinhar à esquerda">⯇</button>
        <button style={sBotao(celSel.al === "center")} onClick={() => formatar({ al: celSel.al === "center" ? undefined : "center" })} title="Centralizar">≡</button>
        <button style={sBotao(celSel.al === "right")} onClick={() => formatar({ al: celSel.al === "right" ? undefined : "right" })} title="Alinhar à direita">⯈</button>
        <button style={sBotao(celSel.qb)} onClick={() => alternar("qb")} title="Quebrar texto automaticamente (Alt+Enter cria nova linha na célula)">↵ Quebrar</button>
        <button style={sBotao(!!mesclDe(sel.r, sel.c))} onClick={mesclar} title="Mesclar e centralizar / desfazer mesclagem">⊞ Mesclar</button>
        <select style={{ ...sSelect, width: 92 }} value="" onChange={(e) => { bordas(e.target.value); e.target.value = ""; }} title="Bordas">
          <option value="" disabled>▦ Bordas</option>
          <option value="todas">Todas as bordas</option>
          <option value="externas">Bordas externas</option>
          <option value="inferior">Borda inferior</option>
          <option value="superior">Borda superior</option>
          <option value="esquerda">Borda esquerda</option>
          <option value="direita">Borda direita</option>
          <option value="nenhuma">Sem borda</option>
        </select>
        <span style={sSep} />
        <select style={sSelect} value={celSel.fmt || ""} onChange={(e) => formatar({ fmt: e.target.value || undefined })} title="Formato do número">
          <option value="">Geral</option>
          <option value="numero">Número (0,00)</option>
          <option value="inteiro">Inteiro</option>
          <option value="moeda">Moeda (R$)</option>
          <option value="pct">Porcentagem</option>
          <option value="data">Data</option>
        </select>
        <button style={sBotao()} onClick={autoSoma} title="AutoSoma">Σ</button>
        <span style={sSep} />
        <button style={sBotao()} onClick={() => estrutura("r", 1)} title="Inserir linha acima">+ Linha</button>
        <button style={sBotao()} onClick={() => estrutura("r", -1)} title="Excluir linha">− Linha</button>
        <button style={sBotao()} onClick={() => estrutura("c", 1)} title="Inserir coluna à esquerda">+ Coluna</button>
        <button style={sBotao()} onClick={() => estrutura("c", -1)} title="Excluir coluna">− Coluna</button>
        <span style={sSep} />
        <button style={sBotao()} onClick={() => ordenar(false)} title="Ordenar de A a Z (menor para o maior) pela coluna selecionada">A→Z</button>
        <button style={sBotao()} onClick={() => ordenar(true)} title="Ordenar de Z a A (maior para o menor) pela coluna selecionada">Z→A</button>
        <select style={{ ...sSelect, width: 104, background: cong.r || cong.c ? "rgba(154,110,30,0.2)" : "#fff" }} value="" onChange={(e) => { congelar(e.target.value); e.target.value = ""; }} title="Congelar painéis">
          <option value="" disabled>❄ Congelar</option>
          <option value="paineis">Congelar painéis (acima e à esquerda da célula)</option>
          <option value="linha">Congelar linha superior</option>
          <option value="coluna">Congelar primeira coluna</option>
          <option value="nenhum">Descongelar</option>
        </select>
        <button style={sBotao(!!filtro)} onClick={alternarFiltro} title={filtro ? "Remover filtro" : "Filtrar a tabela da célula selecionada"}>⏷ Filtro</button>
      </div>

      <div style={{ display: "flex", gap: 4, alignItems: "center", padding: "4px 6px", background: "#fff", borderLeft: "1px solid rgba(0,0,0,0.12)", borderRight: "1px solid rgba(0,0,0,0.12)" }}>
        <input value={nomeCaixa ?? refSel} onChange={(e) => setNomeCaixa(e.target.value)} onFocus={(e) => e.target.select()}
          onKeyDown={(e) => { if (e.key === "Enter") irPara(e.target.value); if (e.key === "Escape") { setNomeCaixa(null); grade.current?.focus(); } }}
          onBlur={() => setNomeCaixa(null)}
          style={{ ...sSelect, width: 90, textAlign: "center", fontWeight: 500 }} title="Ir para célula" />
        <span style={{ color: "#888", fontStyle: "italic", fontFamily: "Georgia, serif", padding: "0 4px" }}>fx</span>
        <input value={textoBarra}
          onFocus={() => { if (!editando) comecarEdicao(celSel.e || "", false, "barra"); }}
          onChange={(e) => setRascunho(e.target.value)}
          onKeyDown={teclaEdicao}
          onBlur={() => { if (editando === "barra") confirmar(); }}
          style={{ ...sSelect, flex: 1, padding: "0 8px", fontFamily: "Consolas, monospace" }} />
      </div>

      <div ref={grade} tabIndex={0} onKeyDown={teclaGrade} onCopy={copiar} onCut={(e) => copiar(e, true)} onPaste={colar} onScroll={rolar}
        style={{ overflow: "auto", height: "62vh", border: "1px solid rgba(0,0,0,0.12)", background: "#fff", outline: "none", position: "relative" }}>
        <table style={{ borderCollapse: "separate", borderSpacing: 0, tableLayout: "fixed", fontSize: 13 }}>
          <thead>
            <tr>
              <th data-cab style={{ ...sCab, top: 0, left: 0, zIndex: 5, width: 44, minWidth: 44, height: ALT_LINHA }}
                ref={(el) => { if (el) larguras.current.set("cab", el.offsetWidth); }}
                onMouseDown={(e) => { e.preventDefault(); setAnc({ r: planilha.linhas - 1, c: planilha.colunas - 1 }); setSel({ r: 0, c: 0 }); grade.current?.focus(); }} />
              {colunas.map((c) => {
                const w = planilha.larg?.[c] || LARG_PADRAO;
                return (
                  <th key={c} data-col={c} style={{ ...sCab, top: 0, left: c < cong.c ? esqCong[c] : undefined, zIndex: c < cong.c ? 5 : 3, width: w, minWidth: w, maxWidth: w, background: c >= c1 && c <= c2 ? "#e2d6b8" : sCab.background }}
                    onMouseDown={(e) => selecionarColuna(e, c)}>
                    <div style={{ position: "relative" }}>
                      {nomeColuna(c)}
                      <span onMouseDown={(e) => redimensionar(e, c)} style={{ position: "absolute", right: -3, top: -4, bottom: -4, width: 7, cursor: "col-resize" }} />
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {linhasDesenho.map((r, i) => (
              <Fragment key={r}>
                {i === nCong && espacoAntes > 0 && <tr><td colSpan={planilha.colunas + 1} style={{ height: espacoAntes, padding: 0, border: 0 }} /></tr>}
                {ocultas.has(r) ? null : (
                  <tr data-r={r}>
                    <th style={{ ...sCab, left: 0, top: r < cong.r ? topoCong(r) : undefined, zIndex: r < cong.r ? 4 : 2, height: planilha.alt?.[r] || ALT_LINHA, background: r >= r1 && r <= r2 ? "#e2d6b8" : sCab.background, color: ocultas.size && filtro && r > filtro.r ? "#1f5fa8" : sCab.color }}
                      onMouseDown={(e) => selecionarLinha(e, r)}>
                      {r + 1}
                      <span onMouseDown={(e) => redimensionarLinha(e, r)} style={{ position: "absolute", left: 0, right: 0, bottom: -3, height: 6, cursor: "row-resize" }} />
                    </th>
                    {colunas.map((c) => celula(r, c))}
                  </tr>
                )}
              </Fragment>
            ))}
            {espacoDepois > 0 && <tr><td colSpan={planilha.colunas + 1} style={{ height: espacoDepois, padding: 0, border: 0 }} /></tr>}
          </tbody>
        </table>
        <div style={{ padding: 8, display: "flex", gap: 8, position: "sticky", left: 0 }}>
          <button style={sBotao()} onClick={() => mudarAba((p) => ({ ...p, linhas: p.linhas + 100 }))}>+ 100 linhas</button>
          <button style={sBotao()} onClick={() => mudarAba((p) => ({ ...p, colunas: p.colunas + 10 }))}>+ 10 colunas</button>
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap", padding: "4px 6px", background: "#f4f1ea", border: "1px solid rgba(0,0,0,0.12)", borderTop: 0, borderRadius: "0 0 10px 10px" }}>
        {pasta.planilhas.map((p, i) => (
          <button key={i} onClick={() => { if (editando) confirmar(); setAba(i); setSel({ r: 0, c: 0 }); setAnc({ r: 0, c: 0 }); }} onDoubleClick={() => renomearAba(i)}
            title="Clique duplo para renomear"
            style={{ ...sBotao(i === abaIdx), borderRadius: "0 0 6px 6px", fontWeight: i === abaIdx ? 600 : 400, color: i === abaIdx ? GOLD : "#444" }}>
            {p.nome}
          </button>
        ))}
        <button style={sBotao()} onClick={adicionarAba} title="Nova planilha">＋</button>
        {pasta.planilhas.length > 1 && <button style={sBotao()} onClick={excluirAba} title="Excluir esta planilha">×</button>}
        {ocultas.size > 0 && <span style={{ fontSize: 12, color: "#1f5fa8", marginLeft: 8 }}>Filtro ativo: {ocultas.size} linha{ocultas.size === 1 ? "" : "s"} oculta{ocultas.size === 1 ? "" : "s"}</span>}
        <span style={{ marginLeft: "auto", fontSize: 12, color: "#555" }}>
          {resumo
            ? <>Soma: <b>{exibir(resumo.soma)}</b> · Média: <b>{resumo.n ? exibir(resumo.soma / resumo.n) : "—"}</b> · Contagem: <b>{resumo.cont}</b></>
            : "Salvo automaticamente neste navegador"}
        </span>
      </div>
      {menuFiltro && <MenuFiltro menu={menuFiltro} setMenu={setMenuFiltro} aplicar={aplicarFiltro}
        ordenar={(desc) => { setSel({ r: filtro.r, c: menuFiltro.c }); setAnc({ r: filtro.r, c: menuFiltro.c }); ordenarLinhas({ r1: filtro.r + 1, r2: fimFiltro(planilha, filtro), c1: filtro.c1, c2: filtro.c2 }, menuFiltro.c, desc); setMenuFiltro(null); }} />}
      <p style={{ fontSize: 12, color: "#777", marginTop: 8, lineHeight: 1.5 }}>
        Fórmulas como no Excel em português: <code>=SOMA(A1:A10)</code>, <code>=MÉDIA(B2:B9)</code>, <code>=SE(C2&gt;100;"Alto";"Baixo")</code>, <code>=PROCV(…)</code>, <code>=SOMASE(…)</code>, <code>=HOJE()</code>.
        Use <b>;</b> entre os argumentos e <b>,</b> nos decimais. A→Z / Z→A ordenam pela coluna selecionada; ⏷ Filtro mostra só as linhas escolhidas. Clique duplo ou F2 edita; Alt+Enter quebra linha na célula; arraste o quadradinho no canto da seleção para preencher séries (1, 2, 3…, datas, meses); Ctrl+D preenche para baixo; Ctrl+C/Ctrl+V copia e cola (inclusive do Excel).
      </p>
    </div>
  );
}

function MenuFiltro({ menu, setMenu, aplicar, ordenar }) {
  const caixa = useRef(null);
  useEffect(() => {
    const fora = (e) => { if (!caixa.current?.contains(e.target)) setMenu(null); };
    const esc = (e) => { if (e.key === "Escape") setMenu(null); };
    window.addEventListener("mousedown", fora);
    window.addEventListener("keydown", esc);
    return () => { window.removeEventListener("mousedown", fora); window.removeEventListener("keydown", esc); };
  }, [setMenu]);
  const busca = menu.busca.trim().toLowerCase();
  const visiveis = menu.itens.filter((t) => !busca || (t || "(Vazias)").toLowerCase().includes(busca));
  const todos = visiveis.every((t) => menu.marcados.has(t));
  const marcar = (lista, ligar) => {
    const m = new Set(menu.marcados);
    lista.forEach((t) => (ligar ? m.add(t) : m.delete(t)));
    setMenu({ ...menu, marcados: m });
  };
  // Com busca, o OK mostra só os itens encontrados que estão marcados (como no Excel).
  const ok = () => aplicar(menu.itens.filter((t) => !menu.marcados.has(t) || (busca && !visiveis.includes(t))));
  const sItem = { display: "flex", alignItems: "center", gap: 6, padding: "3px 4px", fontSize: 13, cursor: "pointer", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" };
  return (
    <div ref={caixa} style={{ position: "fixed", left: Math.max(8, menu.x), top: menu.y, zIndex: 50, width: 260, background: "#fff", border: "1px solid rgba(0,0,0,0.2)", borderRadius: 8, boxShadow: "0 6px 24px rgba(0,0,0,0.2)", padding: 8, fontFamily: "sans-serif" }}>
      <button style={{ ...sBotao(), width: "100%", textAlign: "left", marginBottom: 4 }} onClick={() => ordenar(false)}>A→Z  Ordenar do menor para o maior</button>
      <button style={{ ...sBotao(), width: "100%", textAlign: "left", marginBottom: 8 }} onClick={() => ordenar(true)}>Z→A  Ordenar do maior para o menor</button>
      <input autoFocus placeholder="Pesquisar" value={menu.busca} onChange={(e) => setMenu({ ...menu, busca: e.target.value })}
        onKeyDown={(e) => { if (e.key === "Enter") ok(); }}
        style={{ ...sSelect, width: "100%", boxSizing: "border-box", padding: "0 8px", marginBottom: 6 }} />
      <div style={{ maxHeight: 220, overflowY: "auto", border: "1px solid #eee", borderRadius: 6, padding: 4 }}>
        <label style={{ ...sItem, fontWeight: 600 }}>
          <input type="checkbox" checked={todos} onChange={() => marcar(visiveis, !todos)} /> (Selecionar tudo)
        </label>
        {visiveis.map((t) => (
          <label key={t} style={sItem} title={t || "(Vazias)"}>
            <input type="checkbox" checked={menu.marcados.has(t)} onChange={(e) => marcar([t], e.target.checked)} />
            {t === "" ? <i style={{ color: "#888" }}>(Vazias)</i> : t}
          </label>
        ))}
        {!visiveis.length && <p style={{ fontSize: 12, color: "#888", margin: 4 }}>Nenhum valor encontrado.</p>}
      </div>
      <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
        <button style={{ ...sBotao(), flex: 1 }} onClick={() => aplicar([])}>Limpar filtro</button>
        <button style={{ ...sBotao(), flex: 1 }} onClick={() => setMenu(null)}>Cancelar</button>
        <button style={{ ...sBotao(), flex: 1, background: GOLD, color: "#fff", border: "none" }} onClick={ok}>OK</button>
      </div>
    </div>
  );
}
