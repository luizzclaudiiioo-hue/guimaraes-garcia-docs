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

## Texto e Planilha

Além do gerador de documentos e das publicações, o app tem duas abas de escritório:

- **📝 Texto** (estilo Word): página A4, Ofício ou Carta (retrato ou paisagem) com
  margens ajustáveis; cabeçalho, rodapé e número de página ("Página X de Y");
  marcação de onde cada página termina; entrelinhas, espaço depois do parágrafo e
  recuo da primeira linha; negrito, itálico, fontes, cores, títulos, listas,
  tabelas, imagens e links; localizar e substituir (Ctrl+L / Ctrl+H).
  Abre `.docx` mantendo fontes, espaçamentos, numeração, cabeçalho, rodapé e
  imagens; salva em Word (`.docx`) e imprime ou gera PDF.
- **📊 Planilha** (estilo Excel): várias planilhas por arquivo, fórmulas em português
  (`=SOMA(A1:A10)`, `=SE(...)`, `=PROCV(...)`, `=PROCX(...)`, `=SOMASES(...)`,
  `=DIATRABALHOTOTAL(...)`, `=DATADIF(...)`, `=PGTO(...)` e outras, com `;` entre
  argumentos), formatos de número (R$, %, data), fontes, cores, bordas, mesclar e
  quebrar texto, congelar painéis, alça de preenchimento (séries de números, datas e
  meses), inserir e excluir linhas/colunas, ordenar, filtrar, copiar e colar
  (inclusive do Excel), desfazer e AutoSoma. Abre e salva `.xlsx` e `.csv`.

Os arquivos ficam salvos automaticamente no navegador de quem usa (não vão para
o servidor). Para guardar ou compartilhar, salve em `.docx` ou `.xlsx`.
O motor de fórmulas fica em `src/formulas.js` e a conversão de/para `.docx` em
`src/docxConversao.js`.
