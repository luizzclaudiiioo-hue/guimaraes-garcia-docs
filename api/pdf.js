// Gera o PDF do editor de Texto no servidor, com o Chromium: o navegador do escritório só
// manda a página já pronta (HTML com estilos, fontes, imagens e o montador de páginas
// embutidos) e recebe o arquivo .pdf para baixar, sem a janela de impressão.
import { exigirSessao } from './_auth.js';

let navegador = null;

async function abrirNavegador() {
  if (navegador?.connected) return navegador;
  const puppeteer = (await import('puppeteer-core')).default;
  // Fora da Vercel (testes locais) dá para apontar um Chromium instalado em CHROMIUM_PATH.
  if (process.env.CHROMIUM_PATH) {
    navegador = await puppeteer.launch({ executablePath: process.env.CHROMIUM_PATH, headless: true, args: ['--no-sandbox'] });
  } else {
    const chromium = (await import('@sparticuz/chromium')).default;
    navegador = await puppeteer.launch({ executablePath: await chromium.executablePath(), headless: true, args: chromium.args });
  }
  return navegador;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Use POST.' });
  if (!exigirSessao(req, res)) return;
  const { html } = req.body || {};
  // A página avisa quando está pronta (window.__pdfPronto); sem isso, não há o que gerar.
  if (typeof html !== 'string' || !html.includes('__pdfPronto')) return res.status(400).json({ error: 'Página para o PDF inválida.' });

  let pagina;
  try {
    const nav = await abrirNavegador();
    pagina = await nav.newPage();
    // Tudo já vem embutido: nenhum acesso à rede a partir da página.
    await pagina.setRequestInterception(true);
    pagina.on('request', (r) => (/^(data|about|blob):/.test(r.url()) ? r.continue() : r.abort()));
    await pagina.setJavaScriptEnabled(true);
    await pagina.setContent(html, { waitUntil: 'load', timeout: 30000 });
    await pagina.waitForFunction('window.__pdfPronto === true || typeof window.__pdfErro === "string"', { timeout: 45000 });
    const erro = await pagina.evaluate(() => window.__pdfErro);
    if (erro) throw new Error(erro);
    const pdf = await pagina.pdf({ preferCSSPageSize: true, printBackground: true });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Cache-Control', 'no-store');
    res.status(200).send(Buffer.from(pdf));
  } catch (e) {
    res.status(500).json({ error: 'Não foi possível gerar o PDF: ' + (e?.message || e) });
  } finally {
    await pagina?.close().catch(() => {});
  }
}
