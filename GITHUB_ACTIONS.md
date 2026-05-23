# Rodar a análise no GitHub Actions (sem servidor)

A automação local (Windows Task Scheduler) só funciona com seu PC ligado.
No GitHub Actions, o GitHub roda o script todo dia, de graça, num servidor
deles. Você só recebe o e-mail.

## Pré-requisitos

- Conta no GitHub (gratuita) — https://github.com/signup
- Git instalado (já tem, vem com o Git Bash que você usa)
- O `email_config.json` já preenchido localmente (pra ter as creds em mãos)

## Passo 1 — Criar o repositório PRIVADO no GitHub

1. Abra https://github.com/new
2. Repository name: `findme-dashboard` (ou outro nome)
3. **MUITO IMPORTANTE: marque "Private"** (não Public — vai ter UUIDs e
   estrutura interna que não precisa ser pública)
4. NÃO marque "Add README", "Add .gitignore", "Add license" (já temos)
5. Clica "Create repository"

A próxima tela mostra comandos. Copie a URL do repo, algo como:
`https://github.com/seu-usuario/findme-dashboard.git`

## Passo 2 — Inicializar git localmente e subir o código

Abra um terminal na pasta do projeto e rode:

```bash
cd C:/Users/GAPPE/Documents/findme_dashboard

git init -b main
git add .
git commit -m "primeira versao"
git remote add origin https://github.com/SEU-USUARIO/findme-dashboard.git
git push -u origin main
```

(Troca `SEU-USUARIO` pelo seu username do GitHub.)

Vai pedir login no GitHub na primeira vez. Use seu usuário + um **Personal
Access Token** (não a senha — GitHub não aceita senha mais). Gerar token em:
https://github.com/settings/tokens?type=beta — marca "Repository access:
this repository" e "Contents: read and write".

## Passo 3 — Configurar os 4 Secrets no GitHub

No GitHub, abra seu repo e vá em:
**Settings** → (menu lateral) **Secrets and variables** → **Actions** →
**New repository secret**

Adicione um por um (clica "New repository secret" pra cada):

| Name | Value |
|---|---|
| `FINDME_PASSWORD` | a senha do seu login FindMe |
| `SMTP_USER` | seu Gmail (ex.: `voce@gmail.com`) |
| `SMTP_PASSWORD` | a app password do Gmail (16 chars, sem espaços) |
| `EMAIL_TO` | destinatários separados por vírgula (ex.: `voce@gmail.com,chefe@gmail.com`) |

Os secrets ficam criptografados; nem você consegue ler depois (só editar/apagar).

## Passo 4 — Disparar a primeira execução manualmente (teste)

1. No GitHub, abra seu repo → aba **Actions**
2. À esquerda, clique no workflow **"FindMe — Análise Diária"**
3. À direita, botão **"Run workflow"** → branch `main` → confirma
4. Aguarda ~2-3 minutos
5. Clica na execução pra ver os logs em tempo real

Se tudo deu certo: ✅ verde + e-mail na sua caixa. Se vermelho ❌: clique
nos passos pra ver onde falhou (geralmente é um secret errado ou faltando).

## Passo 5 — Pronto. A partir de amanhã roda sozinho às 07h (BR)

A linha `cron: '0 10 * * *'` no [diario.yml](.github/workflows/diario.yml)
significa "às 10h UTC", que é 07h no horário de Brasília. Pra mudar:

| Horário Brasília (UTC-3) | Cron |
|---|---|
| 06:00 | `0 9 * * *` |
| **07:00** | **`0 10 * * *`** ← padrão |
| 08:00 | `0 11 * * *` |
| 12:00 | `0 15 * * *` |

Edita o arquivo, commit, push — passa a valer no próximo dia.

## Como funciona o histórico

A cada execução, o workflow:
1. Baixa o repositório (incluindo `historico/` com snapshots dos dias
   anteriores)
2. Roda a análise (gera novo snapshot pra hoje)
3. **Commita o snapshot novo de volta no repo** (auto-commit)

Assim, na execução do dia seguinte, todo histórico anterior já está lá.
Os padrões persistentes ("crítico há 3 dias") continuam funcionando como
no PC.

## Atualizar o código depois (ex.: novo posto, ajuste no skill)

Faz as alterações localmente, depois:

```bash
git add .
git commit -m "descrição da mudança"
git push
```

GitHub Actions pega a versão nova automaticamente na próxima execução.

## Custo

- Conta GitHub gratuita: **2.000 minutos/mês** de Actions pra repos privados
- Cada execução nossa: ~2-3 minutos
- 30 dias × 3 min = **90 minutos/mês** (4.5% do limite)
- **Custo real: R$ 0,00**

## Coisas que ficam fora do git (e por quê)

Tá tudo no `.gitignore`:

- `config.json`, `email_config.json` — têm senhas. Ficam SÓ no seu PC e nos
  Secrets do GitHub.
- `audit.jsonl`, `.audit-key` — logs do próprio Claude, não do projeto.
- `relatorios/` — regenerados a cada execução, não precisam viajar.
- `.claude/projects/`, `sessions/`, `tasks/`, `backups/` — estado local.
- `logs/` — locais.

## Voltar atrás (rodar só local de novo)

GitHub Actions e Task Scheduler local podem coexistir — mas se quiser
parar o Actions:
1. GitHub → repo → Settings → Actions → General → **Disable actions**
   (ou apaga o arquivo `.github/workflows/diario.yml` e dá push)

E manter o Task Scheduler local rodando normalmente.
