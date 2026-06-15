# findme-dashboard

Automação de relatórios operacionais do **FindMe** (rondas de vigilância e
atividades de limpeza em condomínios). Todo dia, extrai da API do FindMe o
que aconteceu no dia anterior (D-1), cruza com o que **deveria** ter
acontecido, gera um Excel enriquecido e envia por email.

## O problema que isso resolve

O portal FindMe mostra as atividades, mas não responde bem às perguntas de
gestão do dia-a-dia:

- O que foi **feito**, o que ficou **parcial**, o que foi **perdido** ontem?
- As **atividades avulsas** que combinamos com o cliente foram registradas
  e executadas — ou nem chegaram a ser criadas no sistema?
- Quais postos estão **críticos** e por quê (justificativas reais da equipe)?
- Isso é um dia ruim ou um **padrão persistente**?

## Como funciona (visão geral)

```
                       ┌──────────────────────────────────────────┐
config.json ──────────▶│ findme_programacao.py                    │
(credenciais FindMe)   │  • login na API                          │
                       │  • puxa atividades do dia (programadas)  │
postos/*.json ────────▶│  • puxa avulsas perdidas                 │
(avulsas esperadas     │  • injeta avulsas esperadas que o        │
 por local/dia/vezes)  │    sistema nem criou                     │
                       │  → relatorios/<dia>/GERAL_<dia>.xlsx     │
                       └──────────────────┬───────────────────────┘
                                          ▼
                       ┌──────────────────────────────────────────┐
                       │ skill findme-analyst (.claude/skills/)   │
                       │  • ler_relatorio.py  — extrai e cruza    │
                       │  • enriquecer_atividades.py — colore a   │
                       │    aba Atividades, insere faltantes,     │
                       │    atualiza cabeçalhos por local         │
                       │  • snapshot.py — grava histórico do dia  │
                       │  • aprender_postos.py — aprende modelos  │
                       │    novos que aparecem repetidamente      │
                       └──────────────────┬───────────────────────┘
                                          ▼
                       ┌──────────────────────────────────────────┐
email_config.json ────▶│ analise_diaria.py (orquestrador)         │
(SMTP/Gmail)           │  • roda tudo acima na ordem              │
                       │  • monta email com KPIs, piores postos,  │
                       │    "por que falhou" (justificativas)     │
                       │  • anexa o Excel enriquecido             │
                       └──────────────────────────────────────────┘
```

## Os 5 status de atividade (e o que significam)

| Status | Cor no Excel | Significado |
|---|---|---|
| ✓ Completa | verde | feita integralmente |
| ⚠ Incompleta / c/ Justif. | amarelo | iniciada mas não finalizada |
| ✗ Não iniciada | vermelho claro | criada no sistema, ninguém começou |
| ⊘ Perdida | vermelho forte | a janela de execução passou |
| ❌ Esperada — Não Registrada | vermelho escuro | está no `postos/*.json` mas o FindMe **nem criou** — falha de cadastro, não de execução |

A Capa Executiva traz a tabela **PROGRAMADAS × AVULSAS** com essa quebra
por origem — num olhar você vê se o buraco do dia está nas rondas
programadas ou nas avulsas combinadas.

## Estrutura do repositório

```
findme-dashboard/
├── analise_diaria.py        ← orquestrador diário (é o que o cron chama)
├── findme_programacao.py    ← extrator principal: API → Excel posto-a-posto
├── findme_dashboard.py      ← extrator de KPIs agregados (uso eventual)
├── config.json              ← credenciais FindMe (NÃO vai pro git)
├── config.json.template     ← modelo do config (sem senha)
├── email_config.json        ← SMTP/destinatários (NÃO vai pro git)
├── requirements.txt         ← requests + openpyxl (instalar tb matplotlib)
│
├── postos/                  ← ★ registro manual das avulsas esperadas
│   └── <local>.json         ← por local: postos → atividades → dias/vezes
│
├── relatorios/              ← saída (não versionada): um dir por período
│   └── YYYY-MM-DD_YYYY-MM-DD/
│       ├── GERAL_*.xlsx     ← relatório consolidado enriquecido
│       └── WHATSAPP_*.pdf   ← resumo por local (estilo WhatsApp), anexado ao e-mail
│
├── .claude/skills/findme-analyst/   ← skill de análise (Claude Code)
│   ├── SKILL.md             ← instruções do skill
│   ├── scripts/             ← ler_relatorio, enriquecer, snapshot, aprender
│   ├── historico/           ← snapshots diários por local (versionado)
│   ├── observados/          ← modelos observados (auto-aprendizado)
│   └── references/          ← domínio FindMe (status, armadilhas, schemas)
│
└── vps/                     ← automação no VPS (substituiu GitHub Actions)
    ├── README.md            ← ★ tutorial passo-a-passo de instalação
    ├── run_diario.sh        ← wrapper: lock → git pull → análise → push
    ├── notify_failure.py    ← email de alerta em caso de falha
    ├── findme-diario.service / .timer  ← systemd (07:30 BRT diário)
```

## O formato do `postos/<local>.json`

A API do FindMe **não devolve** a lista de avulsas cadastradas — esse
registro é mantido à mão, um arquivo por local:

```json
{
  "local": "CONDOMÍNIO ATUA MOOCA",
  "postos": [
    {
      "posto": "Ronda",
      "op_tipo": "Vigilante",
      "atividades": [
        {
          "modelo": "RONDA NO HALLS",
          "dias": ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab"],
          "vezes": 1
        }
      ]
    }
  ]
}
```

Regras:
- `modelo` = nome **exato** como aparece no FindMe (acentos são normalizados
  no matching, então `CONDOMINIO` × `CONDOMÍNIO` não é problema).
- `vezes: 0` = atividade **sob demanda** — existe no posto mas não é cobrada
  diariamente (não gera "Esperada — Não Registrada").
- `periodo: "quinzena"` = ainda sem cálculo de paridade — tratada como sob
  demanda (não cobra, mas aceita se aparecer).
- Só cadastre aqui as **avulsas**. As rondas programadas o relatório já traz.

## Como rodar

```bash
# análise de ontem (D-1) + email
python analise_diaria.py

# dia específico
python analise_diaria.py --data 2026-06-11

# tudo menos o envio do email (pra conferir antes)
python analise_diaria.py --data 2026-06-11 --sem-email
```

Pré-requisitos: Python 3.10+, `pip install -r requirements.txt openpyxl
matplotlib`, `config.json` e `email_config.json` preenchidos (use os
templates).

## Automação diária

Roda num **VPS Linux** via systemd timer às **07:30 BRT** (10:30 UTC):
puxa o código do GitHub, roda a análise, envia o email e commita os
snapshots de volta. Falhou em qualquer etapa → email de alerta com o log.

Setup completo: **[vps/README.md](vps/README.md)** (tutorial passo-a-passo,
não assume conhecimento de Linux).

> Histórico: a automação rodou no GitHub Actions até maio/2026
> (removida no commit `fc204ba` — atrasos de cron e limites de runner).

## O skill findme-analyst (Claude Code)

Pra análises interativas ("o que falhou ontem?", "por que o Splendor está
crítico?"), o skill em `.claude/skills/findme-analyst/` lê o Excel, cruza
com `postos/` e o histórico, e enriquece o próprio arquivo. Detalhes em
[SKILL.md](.claude/skills/findme-analyst/SKILL.md).

O histórico de snapshots (`historico/YYYY-MM-DD/<local>.json`) é o que
permite dizer "isso se repete há N dias" em vez de só "isso aconteceu hoje"
— por isso ele é versionado no git.

## Armadilhas conhecidas (leia antes de interpretar números)

1. **0% nem sempre é equipe parada.** Local com 0 OK + dezenas de não-feitas
   pode ser artefato: local duplicado no portal, programação fantasma, ou
   posto que deixou de reportar.
2. **Eficiência ≠ cumprimento.** A eficiência pondera parciais (OK + 0,5 ×
   parcial). O cumprimento do cabeçalho considera só feitas.
3. **Avulsas dependem do `postos/*.json` estar correto.** Registro duplicado
   (dois arquivos pro mesmo condomínio) cobra as esperadas em dobro.
4. **Justificativas importam.** "CHUVA FORTE" em 20 atividades muda a leitura
   do dia inteiro — o email traz as categorias e frases reais da equipe.
