# Guimarães — Gerador de Documentos

App para geração automática de Procurações e Contratos de Honorários.

## Como publicar na Vercel

1. Faça upload desta pasta para um repositório no GitHub
2. Acesse vercel.com e importe o repositório
3. Em "Environment Variables", adicione:
   - Nome: `ANTHROPIC_API_KEY`
   - Valor: sua chave da Anthropic (sk-ant-...)
   - `VITE_SENHA_1` e `VITE_SENHA_2` (ou `SENHA_1` e `SENHA_2`): senhas de acesso ao app
   - Opcional: `SEGREDO_SESSAO`, texto aleatório usado para assinar a sessão de login
4. Clique em Deploy

As senhas são conferidas apenas no servidor (`api/login.js`). Nunca use
`import.meta.env.VITE_SENHA_*` no código de `src/`: o Vite embutiria a senha no site.

## Como rodar localmente

```bash
npm install
npm run dev
```

Acesse http://localhost:5173

## Editor de Texto

Além do gerador de documentos e das publicações, o app tem um editor de texto:

- **📝 Texto** (estilo Word): interface com faixa de opções (Arquivo, Página Inicial,
  Inserir, Layout, Referências, Revisão, Exibir), zoom e fundo escuro em volta da
  página; fonte padrão Arial Narrow 11; página A4, Ofício ou Carta (retrato ou paisagem) com
  margens ajustáveis; cabeçalho, rodapé e número de página ("Página X de Y"),
  com opção de primeira página diferente (timbre só na folha de rosto);
  **Exportar PDF / Imprimir** com as páginas montadas como no Word (cabeçalho e
  rodapé em cada página, 1ª página diferente, "Página X de Y", notas no pé da
  própria página e sumário com os números de página exatos);
  **régua** como a do Word (recuo da 1ª linha, deslocado, esquerdo e direito; clique
  cria tabulação, Tab no texto avança até ela) e **quebras de seção** com páginas em
  paisagem no meio do documento (para quadros e tabelas largas), no .docx e no PDF;
  marcação de onde cada página termina; entrelinhas, espaço depois do parágrafo e
  recuo da primeira linha; negrito, itálico, fontes, cores, títulos, listas
  em vários níveis (Tab / Shift+Tab; formatos 1. 1.1. 1.1.1., 1. a) i), I. A. 1.),
  bordas e sombreamento de parágrafo, pincel de formatação, data por extenso,
  links e localizar e substituir (Ctrl+L / Ctrl+H).
  - **Tabelas**: seletor de tamanho, menu do botão direito (inserir/excluir linhas
    e colunas, mesclar arrastando, dividir, bordas, sombreamento), largura das
    colunas arrastando, Tab entre células.
  - **Imagens**: alças para redimensionar, em linha, centralizada ou com texto ao redor.
  - **Notas de rodapé**, **estilos editáveis** (Normal, Títulos, Citação) e
    **sumário automático**.
  - **Comentários** em balões ao lado da página (responder, resolver, excluir) e
    **controle de alterações** com autor e data (aceitar/rejeitar uma ou todas).
  Abre `.docx` mantendo fontes, espaçamentos, numeração, cabeçalho, rodapé,
  imagens, tabelas, notas, comentários e alterações; salva em Word (`.docx`) e
  imprime ou gera PDF.

**Abrir e salvar no mesmo arquivo:** no Chrome e no Edge (computador), "📂 Abrir"
liga o documento ao arquivo escolhido, e "💾 Salvar" (Ctrl+S) grava
por cima desse mesmo arquivo; "Salvar como…" (Ctrl+Shift+S) cria outro. Na primeira
gravação o navegador pede permissão. Se o arquivo tiver recursos que o app ainda não
mantém (gráficos, caixas de texto, campos de formulário etc.), o app avisa antes
de gravar por cima. Em outros navegadores, "Salvar" baixa um arquivo novo.

O conteúdo também fica guardado automaticamente no navegador de quem usa (não vai
para o servidor). Código: `src/arquivoLocal.js`.

A conversão de/para `.docx` fica em `src/docxConversao.js`; tabelas, estilos e
revisão em `src/textoTabelas.js`, `src/textoEstilos.js` e `src/textoRevisao.js`.
