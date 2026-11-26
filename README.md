# GradeFácil Escolar

Aplicação web para montar horários escolares sem conflitos: cadastre turmas e professores, gere a grade automaticamente e exporte para **PDF** ou **Word (.docx)**.

Os dados ficam salvos no próprio navegador (localStorage) — nada é enviado a servidores.

## Como usar

1. **Cadastrar turmas** — turno, aulas por dia, duração e horário de início.
2. **Professores e aulas** — vincule professores e disciplinas às turmas, com dias e turnos disponíveis.
3. **Grade completa** — gere a grade, ajuste aulas manualmente e exporte.

## Desenvolvimento

Requer Node.js 20+.

```bash
pnpm install
pnpm dev        # servidor de desenvolvimento
pnpm build      # build de produção em dist/
pnpm preview    # visualizar o build
```

## Tecnologias

React 19 · Vite · TypeScript · Tailwind CSS v4 · docx · mammoth · lucide-react
