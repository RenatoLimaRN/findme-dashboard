# Automação diária — FindMe

Roda todo dia às 07h: gera o relatório do dia anterior, enriquece a aba
Atividades, salva snapshot no histórico, e envia por e-mail com o `.xlsx`
anexado.

## Setup (uma vez só)

### 1. Gerar app password do Gmail

A senha normal do Gmail **não funciona** mais para SMTP — precisa de uma
"app password". Pra gerar:

1. Tenha verificação em 2 etapas ativa na sua conta Google
2. Acesse https://myaccount.google.com/apppasswords
3. Crie uma nova, com nome tipo "FindMe Analyst"
4. Copie os 16 caracteres (vem como `abcd efgh ijkl mnop`)

### 2. Configurar credenciais

```bat
copy email_config.json.template email_config.json
```

Abra `email_config.json` e preencha:
- `user` — seu Gmail
- `password` — a app password gerada
- `to` — destinatários (lista, pode ser só você ou incluir gestão)

**Importante:** se o projeto for versionado, adicione `email_config.json`
ao `.gitignore` — não suba a senha pro Git.

### 3. Testar manualmente antes de agendar

```bat
python analise_diaria.py --data 2026-05-15 --sem-email
```

Confere se gerou o relatório e enriqueceu sem erro. Depois testa com envio:

```bat
python analise_diaria.py --data 2026-05-15
```

E confere o e-mail.

### 4. Agendar pra rodar todo dia

```bat
agendar_analise.bat
```

Cria uma tarefa "FindMe Analise Diaria" no Windows Task Scheduler, que roda
todo dia às 07h. Pra mudar o horário, edite o `set HORA=07:00` no `.bat`
e rode de novo.

## Como funciona

Cada execução faz:

1. **Gera** o relatório do dia anterior (chama `findme_programacao.py`)
2. **Lê** o relatório (`ler_relatorio.py` do skill) — extrai KPIs, cruza com
   `postos/*.json`, busca histórico
3. **Enriquece** a aba Atividades in-place (`enriquecer_atividades.py`):
   cores por status, linhas das esperadas não registradas, cabeçalho com
   resumo
4. **Salva** snapshots em `.claude/skills/findme-analyst/historico/`
5. **Monta** o HTML do resumo (KPIs + ranking + padrões do histórico)
6. **Envia** por e-mail com o `.xlsx` anexado

Em caso de erro, tenta enviar um e-mail com o traceback. Log de cada
execução em `logs/YYYY-MM-DD.log`.

## Uso manual (fora do agendamento)

```bat
python analise_diaria.py                    # ontem (default)
python analise_diaria.py --data 2026-05-15  # dia específico
python analise_diaria.py --sem-email        # roda tudo menos envio
python analise_diaria.py --pular-gerar      # usa relatório existente, só enriquece + envia
```

## Onde olhar quando algo der errado

- **Log da execução:** `logs/YYYY-MM-DD.log`
- **Relatório gerado:** `relatorios/<data>_<data>/GERAL_*.xlsx`
- **Dados extraídos:** `.claude/skills/findme-analyst-workspace/dia-<data>/dados.json`
- **Snapshots do dia:** `.claude/skills/findme-analyst/historico/<data>/`

## Pra desagendar

```bat
schtasks /Delete /TN "FindMe Analise Diaria" /F
```
