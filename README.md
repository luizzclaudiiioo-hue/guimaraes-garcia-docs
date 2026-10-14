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
