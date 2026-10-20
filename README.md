# SIPL — Consulta LCM

Interface para consultar o relatório LCM de materiais por código de OPM.

## Como funciona

O usuário informa o código completo da OPM (9 dígitos), seleciona uma unidade ativa e abre o relatório fornecido pelo SIPL. A lista de unidades é carregada do arquivo `public/tabela OPM.xlsx`.

O relatório é servido pelo sistema interno da PMESP. Por isso, o acesso depende de uma conexão autorizada à intranet. A aplicação não mantém cópias dos relatórios.

## Executar localmente

Requisitos: Node.js 22 e npm.

```bash
npm install
npm run dev
```

Para conferir a versão de produção:

```bash
npm run build
npm run preview
```

## Publicação

O workflow em `.github/workflows/deploy.yml` publica a versão da branch `main` no GitHub Pages após o build.

## Cuidados ao alterar

- Manter a validação de códigos de OPM com 9 dígitos.
- Não modificar a URL do relatório interno sem verificar o funcionamento no SIPL.
- Conferir pesquisa, seleção da unidade, limpeza do formulário, abertura do relatório e orientação para salvar PDF.
- Validar o build antes de publicar mudanças.
