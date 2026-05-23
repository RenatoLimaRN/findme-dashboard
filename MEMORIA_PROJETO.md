# Memória do Projeto — FindMe Dashboard

> Arquivo atualizado a cada conversa. Serve como base de conhecimento para continuar o desenvolvimento de onde parou.

---

## 🎯 Objetivo do Projeto

Construir um sistema em Python que consome a **FindMe Dashboard API v2.0** e gera relatórios Excel e futuramente dashboards HTML interativos, automatizados e com envio por e-mail.

---

## 📁 Estrutura de Arquivos

```
C:\Users\GAPPE\Documents\findme_dashboard\
├── findme_dashboard.py      # Script 1 — KPIs: atividades, checklists, justificativas
├── findme_programacao.py    # Script 2 — Grade de atividades por posto (ronda/limpeza)
├── config.json              # Credenciais e UUIDs dos locais (compartilhado)
├── requirements.txt         # Dependências Python
├── MEMORIA_PROJETO.md       # Este arquivo
└── locais_disponiveis.csv   # Gerado no modo listagem
```

---

## 🔐 Credenciais e Ambiente

- **Login URL:** `https://production.api.findme.id/v3/settings/login`
- **Dashboard Base URL:** `https://dashboard-production.findme.id`
- **Autenticação:** JWT via `Authorization: Bearer <token>`
- **Credencial atual:** `rnl.lima.nascimento@gmail.com` (ver config.json)
- **Login testado e confirmado ✅** — token obtido com sucesso na máquina Windows

### Locais Ativos Monitorados (5 locais)

| UUID | Nome | Cliente | Região |
|---|---|---|---|
| `0f76c9e9-e7a7-4fc9-8452-72ca78393d61` | ACADEMIA GAVIÕES ATIBAIA | ACADEMIA GAVIÕES ATIBAIA | INTERIOR |
| `ccbf3d2f-a32f-4f88-9eca-d96dd285be71` | VARANDA POMPÉIA | VARANDA POMPÉIA | ZONA OESTE |
| `6e0081a7-6691-4491-8af4-d6a74c35cdb8` | EDIFICIO GUADELUPE | EDIFICIO GUADELUPE | ZONA OESTE |
| `36a41d06-c7d0-43de-8014-cd2ebb1fde8d` | CONDOMÍNIO VIVAZ PRIME - RIO BONITO | Vivaz Rio Bonito | ZONA SUL |
| `5db33317-d88a-4fda-afc4-db0661b4e81c` | CONDOMÍNIO ATUA MOOCA | Condomínio Atua Mooca | ZONA LESTE |

---

## ⚠️ Aprendizados Técnicos

### Sandbox vs Máquina Local
- A sandbox Linux usada pelo Claude **NÃO tem acesso à internet externa**
- Proxy retorna `403 Forbidden` ao tentar acessar `production.api.findme.id`
- **Conclusão:** scripts que chamam a API FindMe só rodam na máquina local do usuário (Windows)
- Usar sandbox apenas para: sintaxe, lógica local, geração de arquivos sem chamadas HTTP

### config.json
- Arquivo pode ser truncado se escrito via ferramenta de arquivo do Claude em arquivos grandes
- **Solução:** escrever o config.json direto via Python/bash na sandbox para garantir integridade
- JSON precisa ser válido — testar sempre com `python3 -m json.tool config.json`

### Modo Listagem
- Ativado automaticamente quando `"locations": []` está vazio no config.json
- Gera `locais_disponiveis.csv` com todos os locais ativos/inativos + UUIDs
- Permite ao usuário escolher os locais antes de rodar o relatório completo

### Limitações da API /reports/routines/general
- **504 Gateway Timeout** ocorre ao enviar muitos locais em uma única requisição (testado com 24 locais)
- **IncompleteRead** ocorre com `limit` alto (100+) — servidor corta a conexão no meio da resposta
- **Solução definitiva:** buscar **1 local por vez**, limit=50, MAX_RETRIES=4, timeout=120s
- Retry captura: `Timeout`, `ChunkedEncodingError`, `ConnectionError`
- Avulsas executadas aparecem nesse mesmo endpoint com `single=True` — não há endpoint separado para avulsas
- A tela "Atividades avulsas" do portal (configuração por posto) não tem endpoint público na API v2.0

---

## 🗓️ Como Rodar (Máquina Windows)

```bash
# 1. Instalar dependências (uma vez só)
pip install -r requirements.txt

# 2. Modo listagem — descubra os UUIDs disponíveis
# Deixar "locations": [] no config.json e rodar:
python findme_dashboard.py

# 3. Modo relatório — após preencher os UUIDs no config.json
python findme_dashboard.py
# → Digitar data início e fim no terminal (formato YYYY-MM-DD)
# → Gera: findme_YYYY-MM-DD_YYYY-MM-DD.xlsx
```

---

## 📊 Módulos da API Implementados

### Script 1 — findme_dashboard.py (KPIs)

| Módulo | Endpoints utilizados | Aba no Excel |
|---|---|---|
| Filtros | `/filters/locations` | Locais |
| Atividades | `/activities/count/*`, `/activities/efficiency/period/month` | Atividades Resumo + por Mês |
| Checklists | `/checklists/count`, `/checklists/efficiency`, `/checklists/count/items/name/*` | Checklists |
| Justificativas | `/justifications/count/category/*` | Justificativas |
| Relatórios | `/reports/missed-single-activities` | Avulsas Perdidas |

### Script 2 — findme_programacao.py (Grade de Atividades)

| Módulo | Endpoints utilizados | Aba no Excel |
|---|---|---|
| Relatórios paginados | `/reports/routines/general` (paginado, **50/página**, **1 local por vez**) | Ronda, Limpeza, Portaria, Avulsas |
| Filtros | `/filters/locations` | — |

**Campos extraídos do relatório:**
- `station.operation_type` → filtra por tipo: 2=Limpeza, 3=Portaria, 4=Vigilante/Ronda
- `to_be_started_at` / `to_be_finished_until` → horário programado
- `started_at` / `finished_at` → horário real
- `single` → avulsa (True) ou programada (False)
- `status` → 0=Não iniciada, 1=Incompleta, 2=Completa, 4=Incompleta+Justif., 5=Perdida
- `patrol.name` → modelo de atividade

**Abas geradas:**
- Resumo, Grade por Posto (visual: posto × dia × horário), Ronda, Limpeza, Portaria, Avulsas, Outras

---

## 🔜 Próximos Passos Planejados

1. **[✅] Testar script na máquina local do usuário** — funcionou, 5 locais, Excel gerado
2. **[✅] findme_programacao.py estável** — busca por local, paginação robusta com retry
3. **[ ] Validar dados no Excel** — conferir se os valores batem com o portal FindMe
4. **[ ] Adicionar FindMe Score** — seção 11 da API, KPI mais estratégico
5. **[ ] Adicionar Sempre Alertas** — contagem + eficiência + intervalo de atendimento
6. **[ ] Dashboard HTML** — gráficos interativos com Chart.js, abre no navegador
7. **[ ] Automação** — agendamento via Windows Task Scheduler + envio por e-mail (SMTP)

---

## 💡 Decisões de Design

- **Saída:** Excel (.xlsx) com abas separadas por módulo
- **Filtro:** por `locations` (UUIDs) + `period` (data início/fim)
- **Seleção de locais:** via config.json (não interativo, mais prático para automação)
- **Modo listagem:** ativado com `locations: []` — gera CSV auxiliar
- **Cores:** paleta azul escuro/claro (padrão FindMe)
- **Fórmulas Excel:** usadas para totais e médias (dinâmico)

---

## 📅 Histórico de Sessões

| Data | O que foi feito |
|---|---|
| 2026-05-06 | Leitura e análise completa da API v2.0 |
| 2026-05-06 | Criação do script Python com 6 abas Excel |
| 2026-05-06 | Adição do modo listagem de locais |
| 2026-05-06 | Descoberta: sandbox sem acesso à internet — rodar localmente |
| 2026-05-06 | Criação do arquivo MEMORIA_PROJETO.md |
| 2026-05-07 | ✅ Primeiro teste real na máquina Windows — script rodou com sucesso, 5 locais, Excel gerado |
| 2026-05-07 | Criado findme_programacao.py — grade de atividades por posto (ronda/limpeza/avulsas/dias/horários) |
| 2026-05-07 | Corrigido bug `style_cell(alt=)` no build_aba_grade |
| 2026-05-07 | Investigados endpoints de avulsas — confirmado que não há endpoint separado, avulsas estão no routines com single=True |
| 2026-05-07 | Corrigido IncompleteRead: limit 100→50, retry em ChunkedEncodingError/ConnectionError |
| 2026-05-07 | Corrigido 504 com muitos locais: busca agora é feita 1 local por vez e resultados combinados |
