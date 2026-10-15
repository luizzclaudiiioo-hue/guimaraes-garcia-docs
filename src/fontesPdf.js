// Recursos para o PDF gerado no servidor (api/pdf.js): arquivos do app em base64 e fontes
// de mesmas medidas para as famílias que o servidor não tem. Usado pelo editor de Texto e
// pela prévia dos documentos do gerador.

// Arquivos do app em texto/base64 (para embutir na página do PDF), guardados após a 1ª vez.
const cacheRecursos = new Map();
export function recursoTexto(url) {
  if (!cacheRecursos.has(url)) cacheRecursos.set(url, fetch(url).then((r) => { if (!r.ok) throw new Error("recurso indisponível"); return r.text(); }));
  return cacheRecursos.get(url);
}
export function recursoDataUrl(url) {
  const chave = "data:" + url;
  if (!cacheRecursos.has(chave)) cacheRecursos.set(chave, fetch(url).then((r) => r.blob()).then((b) => new Promise((ok, erro) => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.onerror = erro; fr.readAsDataURL(b); })));
  return cacheRecursos.get(chave);
}
// Fontes de mesmas medidas para o PDF gerado no servidor, que não tem Times, Arial etc.:
// entram só as das famílias usadas no documento.
const ARQ_SUBSTITUTAS = {
    "tinos-400-normal": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-400-normal.woff2", import.meta.url).href,
    "tinos-400-italic": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-400-italic.woff2", import.meta.url).href,
    "tinos-700-normal": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-700-normal.woff2", import.meta.url).href,
    "tinos-700-italic": new URL("../node_modules/@fontsource/tinos/files/tinos-latin-700-italic.woff2", import.meta.url).href,
    "arimo-400-normal": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-400-normal.woff2", import.meta.url).href,
    "arimo-400-italic": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-400-italic.woff2", import.meta.url).href,
    "arimo-700-normal": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-700-normal.woff2", import.meta.url).href,
    "arimo-700-italic": new URL("../node_modules/@fontsource/arimo/files/arimo-latin-700-italic.woff2", import.meta.url).href,
    "carlito-400-normal": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-400-normal.woff2", import.meta.url).href,
    "carlito-400-italic": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-400-italic.woff2", import.meta.url).href,
    "carlito-700-normal": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-700-normal.woff2", import.meta.url).href,
    "carlito-700-italic": new URL("../node_modules/@fontsource/carlito/files/carlito-latin-700-italic.woff2", import.meta.url).href,
    "cousine-400-normal": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-400-normal.woff2", import.meta.url).href,
    "cousine-400-italic": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-400-italic.woff2", import.meta.url).href,
    "cousine-700-normal": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-700-normal.woff2", import.meta.url).href,
    "cousine-700-italic": new URL("../node_modules/@fontsource/cousine/files/cousine-latin-700-italic.woff2", import.meta.url).href,
};
const SUBSTITUTAS = [
  [["Times New Roman", "Times", "Georgia", "Cambria", "Book Antiqua", "Garamond"], "tinos"],
  [["Arial", "Helvetica", "Verdana", "Tahoma", "Segoe UI", "Aptos", "Aptos Display"], "arimo"],
  [["Calibri"], "carlito"],
  [["Courier New", "Courier"], "cousine"],
];
export async function fontesSubstitutas(html) {
  let css = "";
  for (const [familias, arq] of SUBSTITUTAS) {
    const usadas = familias.filter((f) => new RegExp(`["'\\s,:=]${f}["'\\s,;]`, "i").test(html));
    for (const f of usadas) {
      for (const [peso, estilo] of [[400, "normal"], [400, "italic"], [700, "normal"], [700, "italic"]]) {
        const dado = await recursoDataUrl(ARQ_SUBSTITUTAS[`${arq}-${peso}-${estilo}`]);
        css += `@font-face { font-family: "${f}"; font-weight: ${peso}; font-style: ${estilo}; src: local("${f}"), url("${dado}") format("woff2"); }\n`;
      }
    }
  }
  return css;
}
