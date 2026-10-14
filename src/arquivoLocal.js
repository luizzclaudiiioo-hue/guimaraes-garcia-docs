// Abrir e salvar no MESMO arquivo do computador (como no Word/Excel), usando a
// File System Access API do Chrome e do Edge. Nos outros navegadores o app volta
// a abrir uma cópia e salvar baixando um arquivo novo.

export const suportaArquivoLocal = typeof window !== "undefined" && "showOpenFilePicker" in window && "showSaveFilePicker" in window;

export const TIPOS = {
  docx: { description: "Documento do Word", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } },
  texto: { description: "Texto ou página", accept: { "text/plain": [".txt", ".md"], "text/html": [".html", ".htm"] } },
};

const cancelado = (e) => e?.name === "AbortError";

// Retorna { handle, file } ou null se a pessoa cancelar.
export async function escolherArquivo(tipos) {
  try {
    const [handle] = await window.showOpenFilePicker({ types: tipos, excludeAcceptAllOption: false, multiple: false });
    return { handle, file: await handle.getFile() };
  } catch (e) {
    if (cancelado(e)) return null;
    throw e;
  }
}

// Pede permissão de escrita (o navegador pergunta na primeira vez) e grava por cima.
export async function gravarNoArquivo(handle, blob) {
  const opcoes = { mode: "readwrite" };
  const atual = handle.queryPermission ? await handle.queryPermission(opcoes) : "granted";
  if (atual !== "granted" && (handle.requestPermission ? await handle.requestPermission(opcoes) : "granted") !== "granted") {
    throw new Error("Sem permissão para gravar no arquivo. Use \"Salvar como…\" ou permita a edição quando o navegador perguntar.");
  }
  let w;
  try {
    w = await handle.createWritable();
    await w.write(blob);
    await w.close();
  } catch (e) {
    try { await w?.abort(); } catch { /* ignora */ }
    if (["NoModificationAllowedError", "InvalidStateError", "NotReadableError"].includes(e?.name)) {
      throw new Error(`Não foi possível gravar em "${handle.name}". Se ele estiver aberto no Word/Excel ou em outro programa, feche-o e tente de novo.`);
    }
    throw e;
  }
}

// "Salvar como…": retorna o novo handle, ou null se cancelar.
export async function salvarComo(blob, nomeSugerido, tipos) {
  try {
    const handle = await window.showSaveFilePicker({ suggestedName: nomeSugerido, types: tipos });
    await gravarNoArquivo(handle, blob);
    return handle;
  } catch (e) {
    if (cancelado(e)) return null;
    throw e;
  }
}

export function baixar(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nome;
  a.style.display = "none";
  document.body.appendChild(a); // fora da página alguns navegadores ignoram o nome do arquivo
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// ---------- Lembrar qual arquivo cada documento usa (sobrevive a recarregar a página) ----------

function banco() {
  return new Promise((ok, erro) => {
    const req = indexedDB.open("arquivosLocais", 1);
    req.onupgradeneeded = () => req.result.createObjectStore("handles");
    req.onsuccess = () => ok(req.result);
    req.onerror = () => erro(req.error);
  });
}
async function operar(modo, fn) {
  try {
    const db = await banco();
    return await new Promise((ok, erro) => {
      const tx = db.transaction("handles", modo);
      const req = fn(tx.objectStore("handles"));
      tx.oncomplete = () => ok(req?.result);
      tx.onerror = () => erro(tx.error);
    });
  } catch {
    return undefined; // IndexedDB indisponível (aba anônima etc.): só perde a lembrança
  }
}
export const lembrarArquivo = (id, handle) => operar("readwrite", (s) => (handle ? s.put(handle, id) : s.delete(id)));
export const arquivoLembrado = (id) => operar("readonly", (s) => s.get(id));

// ---------- O que o editor não preserva ao salvar por cima ----------

export async function recursosNaoSuportados(buffer, tipo = "docx") {
  try {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(buffer);
    const nomes = Object.keys(zip.files);
    const ler = async (re) => (await Promise.all(nomes.filter((n) => re.test(n)).map((n) => zip.file(n).async("string")))).join("");
    const avisos = [];
    if (tipo === "docx") {
      const doc = await ler(/^word\/document\.xml$/);
      if (nomes.some((n) => /^word\/comments/.test(n))) avisos.push("comentários");
      if (/<w:(ins|del) /.test(doc)) avisos.push("controle de alterações");
      if (nomes.some((n) => /^word\/charts\//.test(n))) avisos.push("gráficos");
      if (/<w:txbxContent/.test(doc)) avisos.push("caixas de texto");
      if (/<w:footnoteReference/.test(doc)) avisos.push("notas de rodapé");
      if (/<w:fldSimple|<w:instrText[^>]*>\s*(TOC|REF|SEQ)/.test(doc)) avisos.push("campos automáticos (sumário, referências)");
      if (/<w:sdt>/.test(doc)) avisos.push("campos de formulário");
    }
    return avisos;
  } catch {
    return [];
  }
}
