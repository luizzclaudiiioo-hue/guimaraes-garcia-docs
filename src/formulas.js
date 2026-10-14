// Motor de fórmulas da Planilha: lê fórmulas no padrão do Excel em português
// (=SOMA(A1:A3;2,5)) e também em inglês (=SUM(A1:A3,2.5)), calcula e converte
// entre os dois padrões (o .xlsx guarda sempre a versão em inglês).

export class Erro {
  constructor(codigo) { this.codigo = codigo; }
  toString() { return this.codigo; }
}
const E = {
  div0: new Erro("#DIV/0!"),
  valor: new Erro("#VALOR!"),
  ref: new Erro("#REF!"),
  nome: new Erro("#NOME?"),
  nd: new Erro("#N/D"),
  num: new Erro("#NÚM!"),
  ciclo: new Erro("#CICLO!"),
};
export const ERROS = E;
const ERRO_EN = { "#DIV/0!": "#DIV/0!", "#VALOR!": "#VALUE!", "#REF!": "#REF!", "#NOME?": "#NAME?", "#N/D": "#N/A", "#NÚM!": "#NUM!", "#CICLO!": "#REF!" };
export const erroEmIngles = (e) => ERRO_EN[e.codigo] || e.codigo;
export function erroDoExcel(codigo) {
  const pt = Object.keys(ERRO_EN).find((k) => ERRO_EN[k] === codigo);
  return new Erro(pt || codigo);
}

// ---------- Endereços ----------

export function nomeColuna(i) {
  let s = "";
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}
export function indiceColuna(nome) {
  let n = 0;
  for (const ch of nome.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export const chave = (r, c) => nomeColuna(c) + (r + 1);
export function lerChave(k) {
  const m = /^([A-Z]+)(\d+)$/.exec(k);
  return m ? { r: Number(m[2]) - 1, c: indiceColuna(m[1]) } : null;
}

// ---------- Datas (número de série do Excel) ----------

const BASE = Date.UTC(1899, 11, 30);
export const dataParaSerie = (a, m, d) => (Date.UTC(a, m - 1, d) - BASE) / 86400000;
export function serieParaData(n) {
  const dt = new Date(BASE + Math.round(n * 86400000));
  return { a: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}
function hoje() {
  const d = new Date();
  return dataParaSerie(d.getFullYear(), d.getMonth() + 1, d.getDate());
}

// ---------- Entrada digitada numa célula ----------

// Converte o texto digitado (sem "=") em valor: número em formato brasileiro,
// moeda, porcentagem ou data. Retorna { v, fmt } onde fmt é o formato sugerido.
export function interpretarEntrada(bruto) {
  const s = String(bruto).trim();
  if (s === "") return { v: null };
  const up = s.toUpperCase();
  if (up === "VERDADEIRO" || up === "TRUE") return { v: true };
  if (up === "FALSO" || up === "FALSE") return { v: false };
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s);
  if (m) {
    const ano = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
    const dia = Number(m[1]), mes = Number(m[2]);
    if (mes >= 1 && mes <= 12 && dia >= 1 && dia <= 31) return { v: dataParaSerie(ano, mes, dia), fmt: "data" };
  }
  m = /^(-?)\s*(R\$)?\s*(-?)([\d.]*\d(?:,\d+)?|\d*,\d+)\s*(%?)$/i.exec(s);
  if (m) {
    let t = m[4];
    // "1.234" com ponto de milhar; "1.5" (um único ponto sem 3 dígitos após) vale como decimal.
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t = t.replace(/\./g, "");
    t = t.replace(",", ".");
    let n = Number(t);
    if (!Number.isFinite(n)) return { v: s };
    if (m[1] || m[3]) n = -n;
    if (m[5]) return { v: n / 100, fmt: "pct" };
    return { v: n, fmt: m[2] ? "moeda" : undefined };
  }
  return { v: s };
}

// ---------- Tokenizador ----------

const RE_REF = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])/;
const RE_NOME = /^[A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*/;

// pt = true: ";" separa argumentos e "," é decimal (com "," fora de número também separando).
// pt = false: "," separa argumentos e "." é decimal.
export function tokenizar(texto, pt = true) {
  const tk = [];
  let i = 0;
  let planilha = null;
  while (i < texto.length) {
    const resto = texto.slice(i);
    const ch = texto[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let j = i + 1, s = "";
      while (j < texto.length) {
        if (texto[j] === '"') { if (texto[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += texto[j++];
      }
      tk.push({ t: "str", v: s });
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      const j = texto.indexOf("'!", i + 1);
      if (j > 0) { planilha = texto.slice(i + 1, j); i = j + 2; continue; }
    }
    if (ch === "#") {
      const m = /^#[A-ZÀ-Ú/0!?]+[!?]?/i.exec(resto);
      if (m) { tk.push({ t: "erro", v: m[0] }); i += m[0].length; continue; }
    }
    const num = pt ? /^\d+(?:[.,]\d+)?(?:[eE][+-]?\d+)?/.exec(resto) : /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(resto);
    if (num) {
      tk.push({ t: "num", v: Number(num[0].replace(",", ".")) });
      i += num[0].length;
      continue;
    }
    const ref = RE_REF.exec(resto);
    if (ref) {
      tk.push({ t: "ref", planilha, absC: !!ref[1], c: indiceColuna(ref[2]), absR: !!ref[3], r: Number(ref[4]) - 1 });
      planilha = null;
      i += ref[0].length;
      continue;
    }
    const nome = RE_NOME.exec(resto);
    if (nome) {
      i += nome[0].length;
      if (texto[i] === "!") { planilha = nome[0]; i++; continue; }
      const up = nome[0].toUpperCase();
      if (texto[i] === "(") tk.push({ t: "func", v: up });
      else if (up === "VERDADEIRO" || up === "TRUE") tk.push({ t: "bool", v: true });
      else if (up === "FALSO" || up === "FALSE") tk.push({ t: "bool", v: false });
      else tk.push({ t: "nome", v: nome[0] });
      continue;
    }
    const op = /^(<>|<=|>=|[-+*/^&=<>%:(),;])/.exec(resto);
    if (op) {
      const o = op[0];
      if (o === ";" || o === ",") tk.push({ t: "sep" });
      else if (o === "(" || o === ")") tk.push({ t: o });
      else tk.push({ t: "op", v: o });
      i += o.length;
      continue;
    }
    throw E.nome;
  }
  return tk;
}

// ---------- Nomes de funções ----------

// português → inglês (o primeiro nome em português de cada função é o usado na exibição)
const PT_EN = {
  "SOMA": "SUM", "MÉDIA": "AVERAGE", "MÍNIMO": "MIN", "MÁXIMO": "MAX", "CONT.NÚM": "COUNT",
  "CONT.VALORES": "COUNTA", "CONTAR.VAZIO": "COUNTBLANK", "SE": "IF", "SEERRO": "IFERROR", "E": "AND", "OU": "OR", "NÃO": "NOT",
  "ARRED": "ROUND", "ARREDONDAR.PARA.CIMA": "ROUNDUP", "ARREDONDAR.PARA.BAIXO": "ROUNDDOWN", "INT": "INT", "ABS": "ABS",
  "MOD": "MOD", "POTÊNCIA": "POWER", "RAIZ": "SQRT", "HOJE": "TODAY", "AGORA": "NOW", "DATA": "DATE",
  "DIA": "DAY", "MÊS": "MONTH", "ANO": "YEAR", "CONCATENAR": "CONCATENATE", "CONCAT": "CONCAT",
  "MAIÚSCULA": "UPPER", "MINÚSCULA": "LOWER", "ARRUMAR": "TRIM", "NÚM.CARACT": "LEN", "ESQUERDA": "LEFT",
  "DIREITA": "RIGHT", "EXT.TEXTO": "MID", "SOMASE": "SUMIF", "CONT.SE": "COUNTIF", "MÉDIASE": "AVERAGEIF",
  "PROCV": "VLOOKUP", "PROCH": "HLOOKUP", "ÍNDICE": "INDEX", "CORRESP": "MATCH", "SOMARPRODUTO": "SUMPRODUCT",
  "PROCX": "XLOOKUP", "SOMASES": "SUMIFS", "CONT.SES": "COUNTIFS", "MÉDIASES": "AVERAGEIFS",
  "DIATRABALHOTOTAL": "NETWORKDAYS", "DIATRABALHO": "WORKDAY", "DATADIF": "DATEDIF", "DIAS": "DAYS",
  "DIA.DA.SEMANA": "WEEKDAY", "FIMMÊS": "EOMONTH", "DATAM": "EDATE", "TEXTO": "TEXT", "VALOR": "VALUE",
  "SUBSTITUIR": "SUBSTITUTE", "MAIOR": "LARGE", "MENOR": "SMALL",
  "VP": "PV", "VF": "FV", "PGTO": "PMT", "TAXA": "RATE", "NPER": "NPER",
  "MED": "MEDIAN", "ÉCÉL.VAZIA": "ISBLANK", "ÉNÚM": "ISNUMBER", "ÉTEXTO": "ISTEXT", "ÉERRO": "ISERROR", "PI": "PI",
};
const semAcento = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "");
const PARA_EN = {};
const PARA_PT = {};
for (const [pt, en] of Object.entries(PT_EN)) {
  PARA_EN[pt] = en;
  PARA_EN[semAcento(pt)] = en;
  PARA_EN[en] = en;
  if (!PARA_PT[en]) PARA_PT[en] = pt;
}
PARA_EN["MEDIA"] = "AVERAGE";
// Funções novas do Excel são gravadas no .xlsx com o prefixo "_xlfn.".
const XLFN = new Set(["XLOOKUP", "CONCAT", "DAYS"]);
export function nomeIngles(nome) {
  nome = nome.replace(/^_XL(FN|WS)\./i, "");
  return PARA_EN[nome] || PARA_EN[semAcento(nome)] || nome;
}
export const nomesFuncoes = () => Object.keys(PT_EN);

// ---------- Reescrita (pt ↔ en, deslocamento de referências) ----------

function textoRef(t) {
  const p = t.planilha ? (/^[A-Za-z_][A-Za-z0-9_]*$/.test(t.planilha) ? t.planilha : `'${t.planilha.replace(/'/g, "''")}'`) + "!" : "";
  if (t.invalida) return p + "#REF!";
  return p + (t.absC ? "$" : "") + nomeColuna(t.c) + (t.absR ? "$" : "") + (t.r + 1);
}

function emitir(tokens, pt) {
  let s = "";
  for (const t of tokens) {
    if (t.t === "num") s += pt ? String(t.v).replace(".", ",") : String(t.v);
    else if (t.t === "str") s += '"' + t.v.replace(/"/g, '""') + '"';
    else if (t.t === "ref") s += textoRef(t);
    else if (t.t === "func") s += pt ? (PARA_PT[nomeIngles(t.v)] || nomeIngles(t.v)) : (XLFN.has(nomeIngles(t.v)) ? "_xlfn." : "") + nomeIngles(t.v);
    else if (t.t === "bool") s += pt ? (t.v ? "VERDADEIRO" : "FALSO") : (t.v ? "TRUE" : "FALSE");
    else if (t.t === "sep") s += pt ? ";" : ",";
    else if (t.t === "erro") s += t.v;
    else if (t.t === "nome") s += t.v;
    else if (t.t === "op") s += t.v;
    else s += t.t;
  }
  return s;
}

// Fórmula digitada (sem "=") → fórmula em inglês para o .xlsx.
export function formulaParaIngles(f) {
  try { return emitir(tokenizar(f, true), false); } catch { return f; }
}
// Fórmula vinda do .xlsx (inglês) → padrão brasileiro.
export function formulaParaPortugues(f) {
  try { return emitir(tokenizar(f, false), true); } catch { return f; }
}

// Aplica fn a cada referência da fórmula (sem "="). fn recebe o token e pode alterá-lo;
// o segundo argumento indica se é início ("ini") ou fim ("fim") de intervalo, ou null.
export function transformarRefs(f, fn) {
  let tk;
  try { tk = tokenizar(f, true); } catch { return f; }
  tk.forEach((t, i) => {
    if (t.t !== "ref") return;
    const papel = tk[i + 1]?.v === ":" ? "ini" : tk[i - 1]?.v === ":" ? "fim" : null;
    fn(t, papel);
    if (t.r < 0 || t.c < 0) t.invalida = true;
  });
  return emitir(tk, true);
}

// Copiar/colar e preencher: desloca referências relativas.
export function deslocarFormula(f, dr, dc) {
  return transformarRefs(f, (t) => {
    if (!t.absR) t.r += dr;
    if (!t.absC) t.c += dc;
  });
}

// ---------- Parser ----------

export function analisar(texto) {
  const tk = tokenizar(texto, true);
  let p = 0;
  const ver = () => tk[p];
  const ehOp = (...ops) => ver()?.t === "op" && ops.includes(ver().v);

  function binario(prox, ops) {
    return () => {
      let esq = prox();
      while (ehOp(...ops)) { const op = tk[p++].v; esq = { t: "bin", op, esq, dir: prox() }; }
      return esq;
    };
  }
  function unario() {
    if (ehOp("-", "+")) { const op = tk[p++].v; const x = unario(); return op === "-" ? { t: "neg", x } : x; }
    return pct();
  }
  function pct() {
    let x = intervalo();
    while (ehOp("%")) { p++; x = { t: "pct", x }; }
    return x;
  }
  function intervalo() {
    const a = primario();
    if (ehOp(":")) {
      p++;
      const b = primario();
      if (a.t !== "ref" || b.t !== "ref") throw E.ref;
      return { t: "int", a, b };
    }
    return a;
  }
  function primario() {
    const t = tk[p++];
    if (!t) throw E.valor;
    if (t.t === "num" || t.t === "str" || t.t === "bool") return { t: "lit", v: t.v };
    if (t.t === "erro") return { t: "lit", v: erroDoExcel(t.v) };
    if (t.t === "ref") return { t: "ref", planilha: t.planilha, r: t.r, c: t.c };
    if (t.t === "(") { const x = comparacao(); if (tk[p++]?.t !== ")") throw E.valor; return x; }
    if (t.t === "func") {
      p++; // "("
      const args = [];
      if (ver()?.t !== ")") {
        for (;;) {
          if (ver()?.t === "sep" || ver()?.t === ")") args.push({ t: "vazio" });
          else args.push(comparacao());
          if (ver()?.t === "sep") { p++; continue; }
          break;
        }
      }
      if (tk[p++]?.t !== ")") throw E.valor;
      return { t: "func", nome: nomeIngles(t.v), args };
    }
    throw E.nome;
  }
  const potencia = binario(unario, ["^"]);
  const mult = binario(potencia, ["*", "/"]);
  const soma = binario(mult, ["+", "-"]);
  const concat = binario(soma, ["&"]);
  const comparacao = binario(concat, ["=", "<>", "<", ">", "<=", ">="]);

  const arvore = comparacao();
  if (p < tk.length) throw E.valor;
  return arvore;
}

// ---------- Avaliação ----------

const ehErro = (v) => v instanceof Erro;

export function paraNumero(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (ehErro(v)) return v;
  const x = interpretarEntrada(v).v;
  return typeof x === "number" ? x : E.valor;
}
export function paraTexto(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "VERDADEIRO" : "FALSO";
  if (typeof v === "number") return String(Number(v.toPrecision(15))).replace(".", ",");
  return String(v);
}
function paraBool(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "string") {
    const u = v.toUpperCase();
    if (u === "VERDADEIRO" || u === "TRUE") return true;
    if (u === "FALSO" || u === "FALSE") return false;
    return E.valor;
  }
  const n = paraNumero(v);
  return ehErro(n) ? n : n !== 0;
}

function comparar(a, b) {
  const tipo = (v) => (v === null ? 0 : typeof v === "number" ? 1 : typeof v === "string" ? 2 : 3);
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ta = tipo(a), tb = tipo(b);
  if (ta !== tb) return ta - tb;
  if (ta === 2) return a.localeCompare(b, "pt-BR", { sensitivity: "base" });
  return a === b ? 0 : a < b ? -1 : 1;
}

// Critério de SOMASE/CONT.SE: ">10", "<>x", "abc*", 5...
function criterio(c) {
  if (typeof c === "number" || typeof c === "boolean") return (v) => comparar(v, c) === 0;
  const m = /^(<>|<=|>=|=|<|>)?(.*)$/.exec(String(c ?? ""));
  const op = m[1] || "=";
  const alvoBruto = m[2];
  const n = interpretarEntrada(alvoBruto).v;
  const alvo = typeof n === "number" ? n : alvoBruto;
  const coringa = typeof alvo === "string" && /[*?]/.test(alvo)
    ? new RegExp("^" + alvo.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i")
    : null;
  return (v) => {
    if (coringa) { const ok = typeof v === "string" && coringa.test(v); return op === "<>" ? !ok : ok; }
    if (typeof alvo === "number" && typeof v !== "number") return op === "<>";
    if (alvo === "" && op === "=") return v === null || v === "";
    const k = comparar(v, alvo);
    return { "=": k === 0, "<>": k !== 0, "<": k < 0, ">": k > 0, "<=": k <= 0, ">=": k >= 0 }[op];
  };
}

class Intervalo {
  constructor(linhas) { this.linhas = linhas; }
  get valores() { return this.linhas.flat(); }
}

function arredondar(n, casas, modo) {
  const f = 10 ** casas;
  const x = Number((Math.abs(n) * f).toPrecision(15));
  const r = modo === "cima" ? Math.ceil(x) : modo === "baixo" ? Math.floor(x) : Math.round(x);
  return (Math.sign(n) * r) / f;
}

// Coleta números dos argumentos: em intervalos ignora texto e vazios (como o Excel).
function numeros(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Intervalo) {
      for (const v of a.valores) { if (ehErro(v)) return v; if (typeof v === "number") out.push(v); }
    } else if (a !== undefined) {
      const n = paraNumero(a);
      if (ehErro(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const FUNCOES = {
  SUM: (a) => { const n = numeros(a); return ehErro(n) ? n : n.reduce((x, y) => x + y, 0); },
  AVERAGE: (a) => { const n = numeros(a); return ehErro(n) ? n : n.length ? n.reduce((x, y) => x + y, 0) / n.length : E.div0; },
  MIN: (a) => { const n = numeros(a); return ehErro(n) ? n : n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = numeros(a); return ehErro(n) ? n : n.length ? Math.max(...n) : 0; },
  MEDIAN: (a) => {
    const n = numeros(a);
    if (ehErro(n)) return n;
    if (!n.length) return E.num;
    n.sort((x, y) => x - y);
    const m = Math.floor(n.length / 2);
    return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
  },
  COUNT: (a) => a.reduce((s, x) => s + (x instanceof Intervalo ? x.valores.filter((v) => typeof v === "number").length : typeof paraNumero(x) === "number" ? 1 : 0), 0),
  COUNTA: (a) => a.reduce((s, x) => s + (x instanceof Intervalo ? x.valores.filter((v) => v !== null && v !== "").length : x !== undefined ? 1 : 0), 0),
  COUNTBLANK: (a) => (a[0] instanceof Intervalo ? a[0].valores.filter((v) => v === null || v === "").length : E.valor),
  SUMPRODUCT: (a) => {
    if (!a.every((x) => x instanceof Intervalo)) return E.valor;
    const listas = a.map((x) => x.valores);
    if (listas.some((l) => l.length !== listas[0].length)) return E.valor;
    return listas[0].reduce((s, _, i) => s + listas.reduce((p, l) => p * (typeof l[i] === "number" ? l[i] : 0), 1), 0);
  },
  ROUND: ([n, c]) => { n = paraNumero(n); c = paraNumero(c ?? 0); return ehErro(n) ? n : ehErro(c) ? c : arredondar(n, Math.trunc(c)); },
  ROUNDUP: ([n, c]) => { n = paraNumero(n); c = paraNumero(c ?? 0); return ehErro(n) ? n : ehErro(c) ? c : arredondar(n, Math.trunc(c), "cima"); },
  ROUNDDOWN: ([n, c]) => { n = paraNumero(n); c = paraNumero(c ?? 0); return ehErro(n) ? n : ehErro(c) ? c : arredondar(n, Math.trunc(c), "baixo"); },
  INT: ([n]) => { n = paraNumero(n); return ehErro(n) ? n : Math.floor(n); },
  ABS: ([n]) => { n = paraNumero(n); return ehErro(n) ? n : Math.abs(n); },
  MOD: ([n, d]) => { n = paraNumero(n); d = paraNumero(d); if (ehErro(n)) return n; if (ehErro(d)) return d; return d === 0 ? E.div0 : n - d * Math.floor(n / d); },
  POWER: ([b, e]) => { b = paraNumero(b); e = paraNumero(e); if (ehErro(b)) return b; if (ehErro(e)) return e; const r = b ** e; return Number.isFinite(r) ? r : E.num; },
  SQRT: ([n]) => { n = paraNumero(n); return ehErro(n) ? n : n < 0 ? E.num : Math.sqrt(n); },
  PI: () => Math.PI,
  TODAY: () => hoje(),
  NOW: () => { const d = new Date(); return hoje() + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
  DATE: ([a, m, d]) => { [a, m, d] = [a, m, d].map(paraNumero); const e = [a, m, d].find(ehErro); return e || (Date.UTC(a, m - 1, d) - BASE) / 86400000; },
  DAY: ([n]) => { n = paraNumero(n); return ehErro(n) ? n : serieParaData(n).d; },
  MONTH: ([n]) => { n = paraNumero(n); return ehErro(n) ? n : serieParaData(n).m; },
  YEAR: ([n]) => { n = paraNumero(n); return ehErro(n) ? n : serieParaData(n).a; },
  AND: (a) => { let r = true; for (const x of a.flatMap((x) => (x instanceof Intervalo ? x.valores.filter((v) => v !== null) : [x]))) { const b = paraBool(x); if (ehErro(b)) return b; r = r && b; } return r; },
  OR: (a) => { let r = false; for (const x of a.flatMap((x) => (x instanceof Intervalo ? x.valores.filter((v) => v !== null) : [x]))) { const b = paraBool(x); if (ehErro(b)) return b; r = r || b; } return r; },
  NOT: ([x]) => { const b = paraBool(x); return ehErro(b) ? b : !b; },
  CONCATENATE: (a) => a.map((x) => (x instanceof Intervalo ? x.valores.map(paraTexto).join("") : paraTexto(x))).join(""),
  CONCAT: (a) => FUNCOES.CONCATENATE(a),
  UPPER: ([s]) => paraTexto(s).toUpperCase(),
  LOWER: ([s]) => paraTexto(s).toLowerCase(),
  TRIM: ([s]) => paraTexto(s).trim().replace(/ +/g, " "),
  LEN: ([s]) => paraTexto(s).length,
  LEFT: ([s, n]) => paraTexto(s).slice(0, n === undefined ? 1 : paraNumero(n)),
  RIGHT: ([s, n]) => { const t = paraTexto(s); const k = n === undefined ? 1 : paraNumero(n); return k > 0 ? t.slice(-k) : ""; },
  MID: ([s, i, n]) => paraTexto(s).substr(paraNumero(i) - 1, paraNumero(n)),
  ISBLANK: ([x]) => x === null,
  ISNUMBER: ([x]) => typeof x === "number",
  ISTEXT: ([x]) => typeof x === "string",
  ISERROR: ([x]) => ehErro(x),
  SUMIF: ([intv, crit, somar]) => {
    if (!(intv instanceof Intervalo)) return E.valor;
    const ok = criterio(crit);
    const alvo = somar instanceof Intervalo ? somar.valores : intv.valores;
    return intv.valores.reduce((s, v, i) => s + (ok(v) && typeof alvo[i] === "number" ? alvo[i] : 0), 0);
  },
  AVERAGEIF: ([intv, crit, med]) => {
    if (!(intv instanceof Intervalo)) return E.valor;
    const ok = criterio(crit);
    const alvo = med instanceof Intervalo ? med.valores : intv.valores;
    const n = intv.valores.map((v, i) => (ok(v) && typeof alvo[i] === "number" ? alvo[i] : null)).filter((x) => x !== null);
    return n.length ? n.reduce((x, y) => x + y, 0) / n.length : E.div0;
  },
  COUNTIF: ([intv, crit]) => (intv instanceof Intervalo ? intv.valores.filter(criterio(crit)).length : E.valor),
  VLOOKUP: ([proc, tab, col, aprox]) => procurar(proc, tab, col, aprox, false),
  HLOOKUP: ([proc, tab, lin, aprox]) => procurar(proc, tab, lin, aprox, true),
  INDEX: ([tab, l, c]) => {
    if (!(tab instanceof Intervalo)) return E.valor;
    let li = paraNumero(l ?? 1), ci = paraNumero(c ?? 1);
    if (tab.linhas.length === 1 && c === undefined) { ci = li; li = 1; }
    const v = tab.linhas[li - 1]?.[ci - 1];
    return v === undefined ? E.ref : v;
  },
  SUMIFS: ([somar, ...pares]) => {
    if (!(somar instanceof Intervalo)) return E.valor;
    const ok = linhasQueAtendem(pares);
    if (ehErro(ok)) return ok;
    const v = somar.valores;
    return ok.reduce((s, i) => s + (typeof v[i] === "number" ? v[i] : 0), 0);
  },
  COUNTIFS: (pares) => { const ok = linhasQueAtendem(pares); return ehErro(ok) ? ok : ok.length; },
  AVERAGEIFS: ([med, ...pares]) => {
    if (!(med instanceof Intervalo)) return E.valor;
    const ok = linhasQueAtendem(pares);
    if (ehErro(ok)) return ok;
    const n = ok.map((i) => med.valores[i]).filter((x) => typeof x === "number");
    return n.length ? n.reduce((x, y) => x + y, 0) / n.length : E.div0;
  },
  XLOOKUP: ([proc, busca, ret, seNao, modo]) => {
    if (!(busca instanceof Intervalo) || !(ret instanceof Intervalo)) return E.valor;
    const vb = busca.valores;
    const m = modo === undefined ? 0 : paraNumero(modo);
    let i = vb.findIndex((v) => comparar(v, proc) === 0);
    if (i < 0 && (m === -1 || m === 1)) {
      let melhor = -1;
      vb.forEach((v, j) => {
        if (v === null || typeof v !== typeof proc) return;
        const k = comparar(v, proc);
        if ((m === -1 && k < 0 && (melhor < 0 || comparar(v, vb[melhor]) > 0)) || (m === 1 && k > 0 && (melhor < 0 || comparar(v, vb[melhor]) < 0))) melhor = j;
      });
      i = melhor;
    }
    if (i < 0) return seNao === undefined ? E.nd : seNao instanceof Intervalo ? E.valor : seNao;
    // Busca em coluna devolve da mesma linha; busca em linha devolve da mesma coluna.
    const vertical = busca.linhas.length > 1 || busca.linhas[0].length === 1;
    const v = vertical ? ret.linhas[i]?.[0] : ret.linhas[0]?.[i];
    return v === undefined ? E.ref : v;
  },
  NETWORKDAYS: ([ini, fim, fer]) => {
    const n = nums([ini, fim], [undefined, undefined]);
    if (ehErro(n)) return n;
    let [a, b] = n.map(Math.floor);
    const sinal = a <= b ? 1 : -1;
    if (sinal < 0) [a, b] = [b, a];
    const f = feriados(fer);
    let total = 0;
    for (let d = a; d <= b; d++) if (diaUtil(d, f)) total++;
    return sinal * total;
  },
  WORKDAY: ([ini, dias, fer]) => {
    const n = nums([ini, dias], [undefined, undefined]);
    if (ehErro(n)) return n;
    let d = Math.floor(n[0]);
    let resto = Math.trunc(n[1]);
    const passo = Math.sign(resto);
    const f = feriados(fer);
    while (resto !== 0) { d += passo; if (diaUtil(d, f)) resto -= passo; }
    return d;
  },
  DATEDIF: ([ini, fim, un]) => {
    const n = nums([ini, fim], [undefined, undefined]);
    if (ehErro(n)) return n;
    const [a, b] = n.map(Math.floor);
    if (a > b) return E.num;
    const x = serieParaData(a), y = serieParaData(b);
    let meses = (y.a - x.a) * 12 + (y.m - x.m) - (y.d < x.d ? 1 : 0);
    switch (paraTexto(un).toUpperCase()) {
      case "D": return b - a;
      case "M": return meses;
      case "Y": return Math.floor(meses / 12);
      case "YM": return meses % 12;
      case "MD": return y.d >= x.d ? y.d - x.d : b - somarMeses(a, meses, false);
      case "YD": { const aniv = somarMeses(a, Math.floor(meses / 12) * 12, false); return b - aniv; }
      default: return E.num;
    }
  },
  DAYS: ([fim, ini]) => { const n = nums([fim, ini], [undefined, undefined]); return ehErro(n) ? n : Math.floor(n[0]) - Math.floor(n[1]); },
  WEEKDAY: ([d, tipo]) => {
    const n = nums([d, tipo], [undefined, 1]);
    if (ehErro(n)) return n;
    const w = diaSemana(n[0]);
    return n[1] === 2 ? (w + 6) % 7 + 1 : n[1] === 3 ? (w + 6) % 7 : w + 1;
  },
  EOMONTH: ([d, m]) => { const n = nums([d, m], [undefined, 0]); return ehErro(n) ? n : somarMeses(Math.floor(n[0]), Math.trunc(n[1]), true); },
  EDATE: ([d, m]) => { const n = nums([d, m], [undefined, 0]); return ehErro(n) ? n : somarMeses(Math.floor(n[0]), Math.trunc(n[1]), false); },
  TEXT: ([v, f]) => formatarTexto(v, f),
  VALUE: ([v]) => paraNumero(v),
  SUBSTITUTE: ([t, velho, novo, qual]) => {
    t = paraTexto(t); velho = paraTexto(velho); novo = paraTexto(novo);
    if (!velho) return t;
    if (qual === undefined) return t.split(velho).join(novo);
    const k = paraNumero(qual);
    let pos = -1;
    for (let i = 0; i < k; i++) { pos = t.indexOf(velho, pos + 1); if (pos < 0) return t; }
    return t.slice(0, pos) + novo + t.slice(pos + velho.length);
  },
  LARGE: ([intv, k]) => { const n = numeros([intv]); if (ehErro(n)) return n; n.sort((a, b) => b - a); const v = n[paraNumero(k) - 1]; return v === undefined ? E.num : v; },
  SMALL: ([intv, k]) => { const n = numeros([intv]); if (ehErro(n)) return n; n.sort((a, b) => a - b); const v = n[paraNumero(k) - 1]; return v === undefined ? E.num : v; },
  PV: (a) => {
    const n = nums(a, [undefined, undefined, undefined, 0, 0]);
    if (ehErro(n)) return n;
    const [r, np, pgto, fv, t] = n;
    if (r === 0) return -(pgto * np + fv);
    const f = (1 + r) ** np;
    return -(pgto * (1 + r * t) * (f - 1) / r + fv) / f;
  },
  FV: (a) => { const n = nums(a, [undefined, undefined, undefined, 0, 0]); return ehErro(n) ? n : vf(n[0], n[1], n[2], n[3], n[4]); },
  PMT: (a) => {
    const n = nums(a, [undefined, undefined, undefined, 0, 0]);
    if (ehErro(n)) return n;
    const [r, np, pv, fv, t] = n;
    if (np === 0) return E.num;
    if (r === 0) return -(pv + fv) / np;
    const f = (1 + r) ** np;
    return -(pv * f + fv) * r / ((1 + r * t) * (f - 1));
  },
  NPER: (a) => {
    const n = nums(a, [undefined, undefined, undefined, 0, 0]);
    if (ehErro(n)) return n;
    const [r, pgto, pv, fv, t] = n;
    if (r === 0) return pgto === 0 ? E.num : -(pv + fv) / pgto;
    const x = (pgto * (1 + r * t) - fv * r) / (pgto * (1 + r * t) + pv * r);
    return x <= 0 ? E.num : Math.log(x) / Math.log(1 + r);
  },
  RATE: (a) => {
    const n = nums(a, [undefined, undefined, undefined, 0, 0, 0.1]);
    if (ehErro(n)) return n;
    const [np, pgto, pv, fv, t, chute] = n;
    const f = (r) => (Math.abs(r) < 1e-12 ? pv + pgto * np + fv : -vf(r, np, pgto, pv, t) + fv);
    let r = chute;
    for (let i = 0; i < 100; i++) {
      const y = f(r);
      const d = (f(r + 1e-7) - y) / 1e-7;
      if (!Number.isFinite(d) || d === 0) break;
      const novo = r - y / d;
      if (Math.abs(novo - r) < 1e-12) return novo;
      r = novo;
    }
    return Math.abs(f(r)) < 1e-6 ? r : E.num;
  },
  MATCH: ([proc, intv, tipo]) => {
    if (!(intv instanceof Intervalo)) return E.nd;
    const vs = intv.valores;
    const t = tipo === undefined ? 1 : paraNumero(tipo);
    if (t === 0) { const i = vs.findIndex((v) => comparar(v, proc) === 0); return i < 0 ? E.nd : i + 1; }
    let achou = -1;
    vs.forEach((v, i) => { if (v !== null && (t > 0 ? comparar(v, proc) <= 0 : comparar(v, proc) >= 0)) achou = i; });
    return achou < 0 ? E.nd : achou + 1;
  },
};

// ----- Condições múltiplas (SOMASES, CONT.SES, MÉDIASES) -----
function linhasQueAtendem(pares) {
  if (pares.length % 2) return E.valor;
  const listas = [];
  for (let i = 0; i < pares.length; i += 2) {
    if (!(pares[i] instanceof Intervalo)) return E.valor;
    listas.push([pares[i].valores, criterio(pares[i + 1])]);
  }
  const n = listas[0][0].length;
  if (listas.some(([v]) => v.length !== n)) return E.valor;
  const ok = [];
  for (let i = 0; i < n; i++) if (listas.every(([v, f]) => f(v[i]))) ok.push(i);
  return ok;
}

// ----- Datas -----
const diaSemana = (n) => new Date(BASE + Math.round(n) * 86400000).getUTCDay(); // 0 = domingo
function feriados(arg) {
  if (arg === undefined) return new Set();
  const lista = arg instanceof Intervalo ? arg.valores : [arg];
  return new Set(lista.filter((v) => typeof v === "number").map(Math.floor));
}
const diaUtil = (n, fer) => { const d = diaSemana(n); return d !== 0 && d !== 6 && !fer.has(n); };
function somarMeses(serie, meses, fimDoMes) {
  const { a, m, d } = serieParaData(serie);
  const ultimo = new Date(Date.UTC(a, m - 1 + meses + 1, 0)).getUTCDate();
  return dataParaSerie(a, m + meses, fimDoMes ? ultimo : Math.min(d, ultimo));
}

// ----- TEXTO(valor; "formato") com os códigos do Excel em português -----
const MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
const SEMANA = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
function formatarTexto(v, fmt) {
  fmt = paraTexto(fmt);
  if (typeof v !== "number") { const n = paraNumero(v); if (ehErro(n)) return paraTexto(v); v = n; }
  if (/[dmay]/i.test(fmt.replace(/"[^"]*"/g, "")) && !/[0#]/.test(fmt)) {
    const { a, m, d } = serieParaData(Math.floor(v));
    const p2 = (x) => String(x).padStart(2, "0");
    return fmt.replace(/"([^"]*)"|dddd|ddd|dd|d|mmmmm|mmmm|mmm|mm|m|aaaa|yyyy|aa|yy/gi, (t, lit) => {
      if (lit !== undefined) return lit;
      switch (t.toLowerCase()) {
        case "dddd": return SEMANA[diaSemana(v)];
        case "ddd": return SEMANA[diaSemana(v)].slice(0, 3);
        case "dd": return p2(d);
        case "d": return String(d);
        case "mmmmm": return MESES[m - 1][0];
        case "mmmm": return MESES[m - 1];
        case "mmm": return MESES[m - 1].slice(0, 3);
        case "mm": return p2(m);
        case "m": return String(m);
        case "aaaa": case "yyyy": return String(a);
        default: return String(a).slice(-2);
      }
    });
  }
  const m = /[#0.,]*[#0][#0.,]*/.exec(fmt);
  if (!m) return fmt;
  const [inteira, dec = ""] = m[0].split(",");
  if (fmt.includes("%")) v *= 100;
  const casas = (dec.match(/[0#]/g) || []).length;
  const numero = new Intl.NumberFormat("pt-BR", {
    minimumFractionDigits: (dec.match(/0/g) || []).length, maximumFractionDigits: casas,
    useGrouping: inteira.includes("."), minimumIntegerDigits: Math.max(1, (inteira.match(/0/g) || []).length),
  }).format(v);
  return (fmt.slice(0, m.index) + numero + fmt.slice(m.index + m[0].length)).replace(/"/g, "");
}

// ----- Financeiras (mesma convenção de sinais do Excel) -----
function vf(r, n, pgto, vp, t) {
  if (r === 0) return -(vp + pgto * n);
  const f = (1 + r) ** n;
  return -(vp * f + pgto * (1 + r * t) * (f - 1) / r);
}
const nums = (args, padroes) => {
  const out = padroes.map((p, i) => (args[i] === undefined ? p : paraNumero(args[i])));
  const e = out.find((x) => ehErro(x) || x === undefined);
  return e === undefined ? out : e || E.valor;
};

function procurar(proc, tab, idx, aprox, horizontal) {
  if (!(tab instanceof Intervalo)) return E.valor;
  const k = paraNumero(idx);
  if (ehErro(k)) return k;
  const linhas = horizontal ? tab.linhas[0].map((_, c) => tab.linhas.map((l) => l[c])) : tab.linhas;
  if (k < 1 || k > linhas[0].length) return E.ref;
  const exato = aprox !== undefined && paraBool(aprox) === false;
  let achou = -1;
  for (let i = 0; i < linhas.length; i++) {
    const c = comparar(linhas[i][0], proc);
    if (c === 0) { achou = i; break; }
    if (!exato && c < 0 && linhas[i][0] !== null) achou = i;
    if (!exato && c > 0) break;
  }
  return achou < 0 ? E.nd : linhas[achou][k - 1];
}

// ctx.celula(planilha, r, c) → valor; ctx.planilha = nome da planilha atual.
export function avaliar(no, ctx) {
  switch (no.t) {
    case "lit": return no.v;
    case "vazio": return null;
    case "ref": return ctx.celula(no.planilha, no.r, no.c);
    case "int": {
      const r1 = Math.min(no.a.r, no.b.r), r2 = Math.max(no.a.r, no.b.r);
      const c1 = Math.min(no.a.c, no.b.c), c2 = Math.max(no.a.c, no.b.c);
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > 200000) return E.ref;
      const linhas = [];
      for (let r = r1; r <= r2; r++) {
        const l = [];
        for (let c = c1; c <= c2; c++) l.push(ctx.celula(no.a.planilha, r, c));
        linhas.push(l);
      }
      return new Intervalo(linhas);
    }
    case "neg": { const n = paraNumero(escalar(avaliar(no.x, ctx))); return ehErro(n) ? n : -n; }
    case "pct": { const n = paraNumero(escalar(avaliar(no.x, ctx))); return ehErro(n) ? n : n / 100; }
    case "bin": {
      const a = escalar(avaliar(no.esq, ctx));
      const b = escalar(avaliar(no.dir, ctx));
      if (ehErro(a)) return a;
      if (ehErro(b)) return b;
      if (no.op === "&") return paraTexto(a) + paraTexto(b);
      if (["=", "<>", "<", ">", "<=", ">="].includes(no.op)) {
        const k = comparar(a, b);
        return { "=": k === 0, "<>": k !== 0, "<": k < 0, ">": k > 0, "<=": k <= 0, ">=": k >= 0 }[no.op];
      }
      const x = paraNumero(a), y = paraNumero(b);
      if (ehErro(x)) return x;
      if (ehErro(y)) return y;
      if (no.op === "+") return x + y;
      if (no.op === "-") return x - y;
      if (no.op === "*") return x * y;
      if (no.op === "/") return y === 0 ? E.div0 : x / y;
      if (no.op === "^") { const r = x ** y; return Number.isFinite(r) ? r : E.num; }
      return E.valor;
    }
    case "func": {
      const nome = no.nome;
      // Funções com avaliação preguiçosa (só calculam o ramo usado)
      if (nome === "IF") {
        const cond = paraBool(escalar(avaliar(no.args[0], ctx)));
        if (ehErro(cond)) return cond;
        const ramo = cond ? no.args[1] : no.args[2];
        if (!ramo) return cond ? true : false;
        return escalar(avaliar(ramo, ctx));
      }
      if (nome === "IFERROR") {
        const v = escalar(avaliar(no.args[0], ctx));
        return ehErro(v) ? escalar(avaliar(no.args[1] || { t: "lit", v: "" }, ctx)) : v;
      }
      const fn = FUNCOES[nome];
      if (!fn) return E.nome;
      const args = no.args.map((a) => (a.t === "vazio" ? undefined : avaliar(a, ctx)));
      const aceitaErro = nome === "ISERROR";
      if (!aceitaErro) { const e = args.find(ehErro); if (e) return e; }
      const r = fn(args);
      return typeof r === "number" && !Number.isFinite(r) ? E.num : r;
    }
    default: return E.valor;
  }
}

// Intervalo em contexto de valor único: só vale se tiver uma célula.
function escalar(v) {
  if (v instanceof Intervalo) return v.linhas.length === 1 && v.linhas[0].length === 1 ? v.linhas[0][0] : E.valor;
  return v;
}
export const valorFinal = escalar;

// ---------- Exibição ----------

const fmtGeral = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 10 });
const fmtNum = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtMoeda = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const fmtPct = new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 2 });
const fmtInt = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });

export function exibir(v, fmt) {
  if (v === null || v === undefined) return "";
  if (ehErro(v)) return v.codigo;
  if (typeof v === "boolean") return v ? "VERDADEIRO" : "FALSO";
  if (typeof v === "number") {
    if (fmt === "moeda") return fmtMoeda.format(v);
    if (fmt === "pct") return fmtPct.format(v);
    if (fmt === "numero") return fmtNum.format(v);
    if (fmt === "inteiro") return fmtInt.format(v);
    if (fmt === "data") {
      const { a, m, d } = serieParaData(v);
      return `${String(d).padStart(2, "0")}/${String(m).padStart(2, "0")}/${a}`;
    }
    if (Math.abs(v) >= 1e15 || (v !== 0 && Math.abs(v) < 1e-9)) return v.toExponential(4).replace(".", ",");
    return fmtGeral.format(Number(v.toPrecision(15)));
  }
  return String(v);
}
