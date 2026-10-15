// Estilos de parágrafo do editor de Texto (como os estilos do Word): Normal,
// Títulos 1 a 3 e Citação. Valem para a tela, a impressão e o .docx.

export const ESTILOS_PADRAO = {
  normal: { nome: "Normal", fonte: "Arial Narrow", tamanho: 11, entrelinhas: 1.5, antes: 0, depois: 6, recuo: 0, alinhamento: "left" },
  h1: { nome: "Título 1", fonte: "Arial Narrow", tamanho: 14, negrito: true, italico: false, cor: "#000000", alinhamento: "left", antes: 12, depois: 6 },
  h2: { nome: "Título 2", fonte: "Arial Narrow", tamanho: 12, negrito: true, italico: false, cor: "#000000", alinhamento: "left", antes: 12, depois: 6 },
  h3: { nome: "Título 3", fonte: "Arial Narrow", tamanho: 11, negrito: true, italico: false, cor: "#000000", alinhamento: "left", antes: 12, depois: 6 },
  citacao: { nome: "Citação", fonte: "Arial Narrow", tamanho: 10, entrelinhas: 1, antes: 0, depois: 6, recuoEsq: 4, alinhamento: "justify", italico: false },
};

// Completa com os padrões (documentos antigos não têm todos os campos).
export function estilosDe(doc) {
  const e = doc?.estilos || {};
  return Object.fromEntries(Object.entries(ESTILOS_PADRAO).map(([k, v]) => [k, { ...v, ...(e[k] || {}) }]));
}

const SERIFADAS = /times|georgia|garamond|cambria|book/i;
const fonte = (f) => `'${String(f).replace(/'/g, "")}', ${SERIFADAS.test(f) ? "serif" : "sans-serif"}`;

export function cssEstilos(estilos, raiz = ".ed-texto") {
  const { normal: n, h1, h2, h3, citacao: c } = estilos;
  const titulo = (sel, t) => `${raiz} ${sel} { font-family: ${fonte(t.fonte)}; font-size: ${t.tamanho}pt; font-weight: ${t.negrito ? "bold" : "normal"}; font-style: ${t.italico ? "italic" : "normal"}; color: ${t.cor}; text-align: ${t.alinhamento}; margin: ${t.antes}pt 0 ${t.depois}pt; line-height: 1.2; text-indent: 0; }`;
  return `
${raiz} { font-family: ${fonte(n.fonte)}; font-size: ${n.tamanho}pt; line-height: ${n.entrelinhas}; }
${raiz} p, ${raiz} div, ${raiz} li { margin: ${n.antes}pt 0 ${n.depois}pt; }
${raiz} p, ${raiz} > div { text-indent: ${n.recuo}cm; text-align: ${n.alinhamento}; }
${titulo("h1", h1)}
${titulo("h2", h2)}
${titulo("h3", h3)}
${raiz} blockquote { font-family: ${fonte(c.fonte)}; font-size: ${c.tamanho}pt; line-height: ${c.entrelinhas}; margin: ${c.antes}pt 0 ${c.depois}pt ${c.recuoEsq}cm; text-align: ${c.alinhamento}; font-style: ${c.italico ? "italic" : "normal"}; }
${raiz} blockquote p { margin: 0; text-indent: 0; text-align: inherit; }
${raiz} td p, ${raiz} td div, ${raiz} li { text-indent: 0; }
`;
}
