# Referência de Domínio — FindMe

Tudo que você precisa saber para diagnosticar um relatório FindMe sem cair nas
armadilhas. Leia por inteiro antes de analisar.

## Índice

1. Os dois tipos de relatório
2. Códigos de status e os três baldes
3. O que cada métrica significa de verdade
4. Armadilhas que invalidam análises
5. Categorias de justificativa e o que indicam
6. Schema do JSON para `escrever_analise.py`

---

## 1. Os dois tipos de relatório

O extrator (`scripts/ler_relatorio.py`) detecta qual é pelo nome das abas. Os
dois são **complementares** — cada um tem algo que o outro não tem.

### Relatório GERAL (gerado pelo `findme_programacao.py`)

Foco: desempenho **posto a posto** + detalhe de cada atividade. Não tem
justificativas nem checklists.

| Aba | Conteúdo |
|---|---|
| `Capa Executiva` | KPIs gerais (% eficiência, locais monitorados, locais críticos, atividades não feitas) + tabela "LOCAIS CRÍTICOS" + tabela "TODOS OS LOCAIS". Colunas das tabelas: `#`, `Local / Posto`, `✓ OK`, `⚠ Parcial`, `✗ Não Feita`, `Total`, `% Efic.`, `Progresso`. |
| `Atividades` | Log linha a linha. Colunas: `Data`, `Hora`, `Turno`, `Tipo / Modelo`, `Status`, `Iniciada`, `Finalizada`, `Duração`, `Pts OK`, `Pts Total`. Agrupada por local (linhas-cabeçalho começam com `📍`). |
| `🏆 Ranking` | Indicadores gerais + TOP 5 melhores locais + TOP 5 que precisam de atenção. |
| `Grade por Posto` | Grade semanal: `Local`, `Posto / Modelo`, `Dom`..`Sáb`. Pode vir vazia. |

### Relatório de KPIs (gerado pelo `findme_dashboard.py`)

Foco: **indicadores agregados da conta inteira** + justificativas + checklists.
Não tem quebra por posto nem detalhe de atividade.

| Aba | Conteúdo |
|---|---|
| `Locais` | `UUID`, `Nome do Local`, `Cliente`, `Região`, `Status`. |
| `Atividades Resumo` | Pares Indicador/Valor: Total de Atividades, Completas, Incompletas, Perdidas, Com Eventos, Não-Conformidades, Check-ins Esperados, Check-ins Realizados, Eficiência (%). |
| `Atividades por Mês` | `Mês`, `Total Atividades`, `Check-ins Esperados`, `Check-ins Feitos`, `Eficiência (%)`. |
| `Checklists` | KPIs (Total de Checklists, Itens Esperados, Itens Verificados, Eficiência %) + "Itens por Nome": `Nome do Item`, `Checklist`, `Ocorrências`, `Tipo`. |
| `Justificativas` | `Justificativa`, `Total Geral`, `Incompletas`, `Perdidas`. |
| `Avulsas Perdidas` | `ID`, `Data da Perda`, `Check-ins Esperados`, `Check-ins Feitos`, `Posto`, `Modelo de Atividade`, `Local`, `Região`, `Cliente`. |

**Se você só tem um dos dois:** trabalhe com o que há, mas diga o que não dá
para concluir. Ex.: com o GERAL você vê *quais* postos falham, mas não *por
quê* (sem justificativas); com o de KPIs você vê as justificativas, mas não
consegue atribuí-las a um posto específico.

---

## 2. Códigos de status e os três baldes

A API usa códigos numéricos; os relatórios usam rótulos. Mapeamento:

| Código | Rótulo | Balde no GERAL |
|---|---|---|
| 2 | Completa | ✓ OK |
| 1 | Incompleta | ⚠ Parcial |
| 4 | Incompleta c/ Justificativa | ⚠ Parcial |
| 0 | Não iniciada | ✗ Não Feita |
| 5 | Perdida | ✗ Não Feita |

No relatório de KPIs, "Incompletas" e "Perdidas" aparecem separadas (não
agrupadas em baldes).

**Taxa de cumprimento** = OK / Total = Completas / Total. É a métrica
operacional mais honesta de "o serviço foi entregue".

Classificação por cumprimento: **Crítico < 70%**, **Atenção 70–90%**,
**OK ≥ 90%**.

---

## 3. O que cada métrica significa de verdade

- **Taxa de cumprimento** (Completas / Total) — % de atividades entregues. É o
  número que importa para "o serviço foi feito?".
- **Eficiência (%)** no relatório de KPIs — é eficiência de **check-in**
  (check-ins realizados / check-ins esperados), **não** a taxa de atividades
  completas. Uma ronda pode estar "Completa" mas com check-ins faltando, e
  vice-versa. São métricas diferentes; sempre deixe claro qual você está
  citando.
- **% Eficiência Geral** na Capa do GERAL — média ponderada de cumprimento dos
  locais. Se vários locais estão com 0% por artefato de dados (ver armadilhas),
  esse número geral fica artificialmente baixo e **não representa a operação
  real**.
- **Atividade avulsa** (prefixo `[AVULSA]` no Tipo/Modelo) — atividade
  não-programada, criada pontualmente. Uma avulsa "Não Feita" tem peso
  diferente de uma ronda programada perdida: pode ser uma demanda que surgiu e
  não foi atendida, ou uma atividade criada e nunca executada.
- **Turno** — Madrugada (00h–06h), Manhã, Tarde, Noite. Útil para ver se as
  falhas se concentram num turno específico (sinal de problema de escala).

---

## 4. Armadilhas que invalidam análises

Estas são as falhas reais que já aconteceram em análises deste projeto. Não
repita:

1. **"Com Eventos" e "Não-Conformidades" NÃO são status.** São marcadores que
   se *sobrepõem* aos status — uma atividade Completa pode ter um evento. Já
   `Completas + Incompletas + Perdidas = Total`. Se você somar "Com Eventos" e
   "Não-Conformes" como se fossem categorias adicionais, o total estoura de
   100%. Cite esses dois como recortes, nunca como fatias do total.

2. **"Perdidas" ≠ "Avulsas Perdidas".** "Perdidas" é a contagem de atividades
   com status Perdida. "Avulsas Perdidas" é um relatório separado de atividades
   avulsas (single) perdidas. Vêm de fontes diferentes da API e podem se
   sobrepor. Se os dois números forem iguais, não é prova de que são a mesma
   coisa — e não trate como dois problemas distintos sem confirmar.

3. **Posto com "tudo zero" é artefato de dados, não falha operacional.** Um
   local com 0 OK, 0 Parcial e centenas/milhares de "Não Feita" quase nunca
   reflete uma equipe que falhou tudo. Reflete um local que não está reportando
   no app, ou uma programação cadastrada que não corresponde à realidade.
   Diagnostique esses postos num grupo à parte ("dados inconclusivos") e
   **não os misture** no ranking de desempenho real — senão a média da operação
   inteira fica distorcida.

4. **Eficiência de check-in vs. taxa de cumprimento.** Ver seção 3. Não chame
   uma de outra. Se o relatório destaca "1.9% de eficiência", verifique se é
   check-in ou cumprimento antes de escrever "a operação cumpriu 1.9%".

5. **Os totais precisam fechar.** Sempre confira `Completas + Incompletas +
   Perdidas = Total` e `% = OK / Total`. Se não fechar, a base tem problema —
   diga isso em vez de seguir.

---

## 5. Categorias de justificativa e o que indicam

As justificativas (só no relatório de KPIs) são a ponte entre *o que* falhou e
*por quê*. As categorias comuns e o que cada uma normalmente aponta:

| Categoria | O que normalmente indica |
|---|---|
| `09 - EMPENHO NO LOCAL` | Efetivo desviado para outra demanda no próprio local. Aponta para **dimensionamento de equipe** — não há gente suficiente para cobrir a atividade e a demanda extra ao mesmo tempo. |
| `01 - ATIVIDADE NÃO FINALIZADA / PERDIDA` | A atividade foi iniciada mas não fechada, ou simplesmente não tocada. Pode ser **processo/app** (não souberam fechar) ou negligência. |
| `04 - APOIO À PORTARIA` | Efetivo desviado para a portaria. Mesmo problema de dimensionamento do `09`, com destino específico. |
| `02 / 06 - ATENDIMENTO AO MORADOR` | Efetivo desviado para uma demanda pontual de morador. Esperado em volume baixo; preocupante se for recorrente. |
| `07 - BATERIA DESCARREGANDO` | Equipamento — o device não foi carregado. Problema **logístico/gestão de equipamento**, totalmente evitável. |
| `08 - CHUVA FORTE` | Ambiental, fora de controle da equipe. Justificativa legítima — não conte contra o desempenho da equipe, mas observe se o volume é alto demais para ser só clima. |
| `10 - VEÍCULO COM PROBLEMAS` | Logística/manutenção de frota. |

Regra geral: justificativas de **efetivo desviado** (`09`, `04`, `02/06`)
apontam para escala/dimensionamento. Justificativas de **equipamento** (`07`,
`10`) apontam para gestão de recursos. **Ambientais** (`08`) são legítimas mas
merecem atenção se forem volume alto. Sempre olhe *qual categoria domina num
posto específico* — a causa-raiz muda o tipo de ação recomendada.

---

## 6. Schema do JSON para `escrever_analise.py`

O renderizador recebe um JSON com uma lista de blocos. Cada bloco tem um
`tipo`. Ordem dos blocos = ordem na aba.

```json
{
  "blocos": [
    { "tipo": "titulo", "texto": "DIAGNÓSTICO OPERACIONAL — <nome/período>" },
    { "tipo": "veredito", "texto": "2 a 4 linhas com o estado real da operação." },
    { "tipo": "secao", "texto": "RANKING DE POSTOS" },
    { "tipo": "paragrafo", "texto": "Texto livre de contexto, se precisar." },
    {
      "tipo": "tabela",
      "colunas": ["Posto", "Total", "Cumprimento", "Classificação"],
      "linhas": [
        ["Mansão do Butantã", 2888, "0%", "Dados inconclusivos"],
        ["Contemporâneo Jardins", 344, "37%", "Crítico"]
      ],
      "realces": ["neutro", "critico"]
    },
    { "tipo": "lista", "itens": ["Item 1 com o porquê.", "Item 2."] }
  ]
}
```

Tipos de bloco:

| `tipo` | Campos | Renderização |
|---|---|---|
| `titulo` | `texto` | Faixa de título no topo da aba. |
| `veredito` | `texto` | Caixa destacada (fundo claro, borda). |
| `secao` | `texto` | Cabeçalho de seção. |
| `paragrafo` | `texto` | Texto corrido. |
| `lista` | `itens` (lista de strings) | Lista com marcadores. |
| `tabela` | `colunas`, `linhas`, `realces` (opcional) | Tabela formatada. |

`realces` é opcional e tem um valor por linha da tabela. Valores aceitos:
`"critico"` (vermelho), `"atencao"` (amarelo), `"ok"` (verde), `"neutro"` ou
`null` (sem cor). Use para colorir o ranking e os postos críticos.

---

## 7. Registro local de atividades esperadas — `postos/*.json`

A API do FindMe **não** devolve a lista de atividades avulsas que estão
cadastradas no sistema (configuração por posto). O usuário mantém essa lista
manualmente, um arquivo por local, em `postos/` na raiz do projeto.

Estrutura (a mesma que o `findme_programacao.py` já consome):

```json
{
  "local": "CONTEMPORÂNEO JARDINS",
  "postos": [
    {
      "posto": "Limpeza",
      "op_tipo": "Limpeza",
      "atividades": [
        { "modelo": "ACADEMIA", "dias": ["Dom","Seg","Ter","Qua","Qui","Sex","Sab"], "vezes": 1 },
        { "modelo": "PISCINA",  "dias": ["Seg","Qua","Sex"], "vezes": 2 }
      ]
    }
  ]
}
```

- `local` — nome do local. Esse é o campo usado para casar com o nome no
  relatório. **Não use o nome do arquivo para casar** — o arquivo só usa o
  slug para organizar.
- `posto` / `op_tipo` — qual posto/tipo (Limpeza, Ronda, Portaria, etc.).
- `atividades[].modelo` — nome exato como aparece no relatório (ex.: "PISCINA",
  "SALÃO DE FESTA"). Match é case-insensitive e sem acento.
- `atividades[].dias` — dias da semana em que aquela atividade deve acontecer.
  Aceitar: "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab" (sem acento).
- `atividades[].vezes` — quantas vezes por dia.

Matching de nome de local entre relatório e `postos/*.json`: case-insensitive,
sem acentos, ignorando o prefixo `CONDOMÍNIO `. Exemplos que casam entre si:
- "CONTEMPORÂNEO JARDINS" ↔ "CONDOMÍNIO CONTEMPORÂNEO JARDINS"
- "MANSÃO DO BUTANTÃ" ↔ "Mansão do Butantã"

Para criar um template vazio para um local novo, use
`scripts/criar_template_posto.py`. Não preencha as `atividades` no chute — só
o usuário sabe o que está cadastrado no portal FindMe.

---

## 8. Cruzamento esperado × feito

A cargo do `ler_relatorio.py` (quando recebe `--data` e `--postos-dir`). Para
cada local com `postos/*.json` correspondente:

1. **Esperadas do dia-alvo** — filtre as `atividades` cujo `dias` contém o dia
   da semana de `--data`. A quantidade esperada é `vezes` (1 por padrão).
2. **Feitas no dia** — agregue, no relatório, as atividades daquele local
   cujo `Data` é igual a `--data` (e/ou prefixo "[AVULSA]" no `Tipo/Modelo`,
   se a programação for avulsa). Conte por `modelo` e por `status` (OK,
   Parcial, Não Feita).
3. **Cruze por modelo** — para cada `modelo` esperado:
   - se `feitas_OK ≥ vezes` → ✓ feita
   - se `feitas_OK + feitas_Parcial ≥ vezes` e `feitas_OK < vezes` → ⚠ parcial
   - se `feitas_Total = 0` → ✗ não feita (não apareceu nem como Não Feita —
     suspeita de programação fantasma OU avulsa esperada que ninguém criou)
   - se `feitas_Total > 0` mas tudo Não Feita → ✗ não feita (apareceu, mas
     nenhuma execução)
4. **Extras** — atividades feitas que não estavam esperadas (modelos que
   apareceram no relatório mas não no registro). Liste-as como "executadas
   além do esperado" — pode ser legítimo (avulsa criada na hora) ou indicar
   registro `postos/*.json` desatualizado.

O JSON de saída deve ter, por local, um campo `cruzamento`:

```json
{
  "esperadas_total": 8,
  "feitas_ok": 5,
  "parciais": 1,
  "perdidas": 2,
  "esperadas_detalhe": [
    {"modelo": "PISCINA", "vezes": 2, "feitas_ok": 1, "parcial": 0, "status": "parcial"},
    {"modelo": "ACADEMIA", "vezes": 1, "feitas_ok": 0, "parcial": 0, "status": "nao_feita"}
  ],
  "extras": [
    {"modelo": "EVENTO ESPECIAL", "total": 1, "ok": 1}
  ]
}
```

Se o local não tem `postos/*.json`, devolva `cruzamento: null` e adicione o
nome ao topo de `postos_sem_registro`.

---

## 9. Histórico de análises — schema do snapshot

Os snapshots vivem em `historico/YYYY-MM-DD/<slug>.json`, um por local por
dia, onde `<slug>` é o nome do local em lowercase com underscores (sem
"CONDOMÍNIO ", sem acentos).

```json
{
  "data": "2026-05-15",
  "local": "CONDOMÍNIO CONTEMPORÂNEO JARDINS",
  "ok": 7,
  "parcial": 2,
  "nao_feita": 3,
  "total": 12,
  "pct_cumprimento": 58.3,
  "dados_inconclusivos": false,
  "esperadas_total": 8,
  "feitas_das_esperadas": 6,
  "perdidas_das_esperadas": 2,
  "modelos_perdidos_hoje": ["PISCINA", "SALÃO DE JOGOS"],
  "top_falhas_modelos": [
    {"modelo": "PISCINA", "nao_feita": 2},
    {"modelo": "SALÃO DE JOGOS", "nao_feita": 1}
  ],
  "obs": ""
}
```

Padrões persistentes que valem a pena chamar na análise (com `historico_recente`
em mãos):

- **Crítico há N dias seguidos** — `pct_cumprimento < 70` em N snapshots
  consecutivos do mesmo local (N ≥ 3 já merece menção).
- **Dados inconclusivos persistentes** — `dados_inconclusivos: true` há N
  dias. Reforço pra escalar para o suporte FindMe.
- **Modelo recorrente em falha** — o mesmo `modelo` aparece em
  `modelos_perdidos_hoje` em 3+ dos últimos 7 snapshots. Aponta para
  problema crônico daquela área/atividade específica.
- **Recuperação** — `pct_cumprimento` subiu de <70 para ≥90 num período curto.
  Vale celebrar.
- **Degradação** — queda contínua mês a mês (se houver série mensal). Algo
  estrutural mudou.

Não force conclusões em cima de 1-2 snapshots — padrão precisa de pelo menos
3 pontos consistentes.

