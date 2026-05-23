---
name: findme-analyst
description: >-
  Analista operacional do FindMe. Foco principal: fechamento do dia anterior
  (D-1) — o que foi feito ontem, o que ficou parcial, o que falhou. Lê
  relatórios Excel do FindMe (findme_dashboard.py KPIs ou findme_programacao.py
  GERAL posto-a-posto), cruza com o registro local de atividades avulsas
  esperadas (postos/*.json), e ENRIQUECE a aba "Atividades" do próprio arquivo:
  colore cada linha por status, adiciona linhas pras avulsas esperadas que nem
  foram registradas no sistema, e atualiza o cabeçalho de cada local com o
  resumo do dia. Mantém histórico em snapshots pra detectar padrões persistentes
  nas próximas rodadas. Use SEMPRE que o usuário pedir pra analisar o dia
  anterior, fechar o dia, ver o que ficou pendente; analisar, diagnosticar,
  interpretar ou "entender" um relatório FindMe; mencionar GERAL_*.xlsx ou
  findme_*.xlsx; perguntar o que foi feito, o que falhou, quais postos estão
  críticos, ou por que os números estão ruins — mesmo sem dizer "análise".
---

# FindMe — Analista Operacional

Pega um relatório FindMe e **enriquece a aba "Atividades" no próprio arquivo**
pra você ver, atividade por atividade, o que foi feito vs. o que ficou pendente
vs. o que sequer chegou a ser registrado no sistema. Foco principal: o
fechamento do dia anterior (D-1).

A análise junta três fontes:

1. **Relatório FindMe** — o que aconteceu (rondas/limpezas feitas, parciais,
   perdidas).
2. **Registro local de atividades esperadas** (em `postos/*.json` na raiz do
   projeto) — o que *deveria* ter acontecido. A API do FindMe não devolve a
   lista de avulsas cadastradas, então o usuário mantém essa lista manualmente.
3. **Histórico de snapshots** (`historico/YYYY-MM-DD/<local>.json` dentro do
   skill) — pra dizer "isso está se repetindo há N dias", não só "isso
   aconteceu hoje".

Sem cruzar (1) com (2) você não consegue dizer se uma avulsa esperada foi feita
ou não. Sem (3) toda análise é amnésica. Não pule essas fontes.

## A mentalidade certa

Um relatório FindMe é cheio de armadilhas:

- **Número ruim nem sempre é operação ruim.** Um posto com 0 OK + 0 Parcial +
  centenas de "Não Feitas" quase sempre é artefato de coleta/cadastro (local
  não está reportando, programação fantasma), não falha da equipe.
- **Causa-raiz, não sintoma.** Ligue as falhas aos modelos, turnos, e
  justificativas.
- **Esperada × executada.** Uma atividade pode estar "Não Feita" porque foi
  perdida (criada no sistema mas não executada) OU pode nem ter sido criada
  (esperada no `postos/*.json` mas o sistema FindMe nem registrou). São duas
  coisas diferentes — a v2.1 distingue: "Não Feita" vs "Esperada — Não
  Registrada".

As armadilhas detalhadas (códigos de status, categorias sobrepostas, eficiência
≠ cumprimento, estrutura dos dois relatórios) estão em
`references/findme-domain.md`. **Leia esse arquivo antes de diagnosticar.**

## Fluxo de trabalho

### 1. Resolva o dia-alvo

Padrão: **ontem (D-1)** — caso de uso principal. Use `--data YYYY-MM-DD` para
outro dia. Se o relatório informado não cobre o dia alvo, avise o usuário e
ofereça gerar um novo via `findme_programacao.py` na raiz do projeto.

### 2. Rode o leitor

```
python scripts/ler_relatorio.py "<arquivo.xlsx>" \
  --data YYYY-MM-DD \
  --postos-dir "<raiz do projeto>/postos" \
  --historico-dir "<skill>/historico" \
  > "<workspace>/dados.json"
```

Saída inclui KPIs, locais, atividades agregadas, **`cruzamento_por_local`** (o
que estava esperado × o que aconteceu), **`historico_por_local`** (últimos 14
dias por local), e **`postos_sem_registro`** (locais sem `postos/*.json`).

### 3. Leia a referência de domínio

`references/findme-domain.md`. Cobre a estrutura dos dois relatórios, semântica
das métricas, armadilhas, justificativas, e os schemas (postos, cruzamento,
snapshot).

### 4. Resolva os postos sem registro

Pra cada local em `postos_sem_registro`, **crie um template vazio**:

```
python scripts/criar_template_posto.py "<raiz do projeto>/postos" "<nome>"
```

Não chute o conteúdo do registro — só o usuário sabe o que está cadastrado no
portal FindMe.

### 5. Enriqueça a aba Atividades

```
python scripts/enriquecer_atividades.py "<arquivo.xlsx>" "<dados.json>"
```

Esse é o passo PRINCIPAL. O script:
- Colore cada linha da aba Atividades por Status (verde=Completa,
  amarelo=Parcial, vermelho=Não Feita/Não iniciada,
  vermelho-escuro=Esperada-Não-Registrada).
- Insere linhas pras avulsas esperadas que nem foram criadas no sistema,
  marcadas como `[ESPERADA — NÃO REGISTRADA] <modelo>` no Tipo/Modelo.
- Atualiza o cabeçalho de cada local (📍) com o resumo do dia:
  `✓ X OK   ⚠ Y Parcial   ✗ Z Não Feitas   ⨯ W Esperadas-não-registradas
  |   N% cumprimento`.

O arquivo é modificado **in-place**. Se preferir não tocar no original,
trabalhe numa cópia.

### 6. Registre snapshots no histórico

Pra cada local com dados reais, salve um snapshot:

```
python scripts/snapshot.py "<skill>/historico" "<snapshot.json>"
```

Schema do snapshot está na referência. **Este passo é o que faz o skill
aprender com o tempo** — sem ele, a próxima análise não detecta padrões
persistentes.

### 7. Responda ao usuário

Resuma em 1-2 parágrafos: como o dia ficou, quais locais merecem atenção, e
quaisquer padrões persistentes que o histórico mostrou. O detalhe fino fica
na aba Atividades que você acabou de enriquecer.

## Convenções de cor na aba Atividades

| Status | Cor de fundo | Significado |
|---|---|---|
| Completa | verde claro | atividade feita integralmente |
| Incompleta / Parcial | amarelo claro | iniciada mas não finalizada |
| Não Feita / Não iniciada / Perdida | vermelho claro | criada no sistema, não executada |
| Esperada — Não Registrada | vermelho-escuro | esperada no `postos/*.json`, mas o FindMe nem criou |

## Arquivos deste skill

- `references/findme-domain.md` — domínio FindMe (estrutura, métricas,
  armadilhas, justificativas, schemas).
- `scripts/ler_relatorio.py` — extrai JSON normalizado do .xlsx + faz o
  cruzamento esperado × feito + consulta o histórico.
- `scripts/enriquecer_atividades.py` — **passo principal**: colore + insere
  esperadas faltantes + atualiza cabeçalhos na aba Atividades, in-place.
- `scripts/snapshot.py` — registra um snapshot diário por local em
  `historico/YYYY-MM-DD/<slug>.json`. Roda ao final de cada análise.
- `scripts/criar_template_posto.py` — cria `postos/<slug>.json` vazio para
  locais sem registro de avulsas.
- `scripts/escrever_analise.py` — DEPRECADO no fluxo atual (criava uma aba
  "Análise" separada). Mantido pra uso futuro caso queira gerar relatórios
  standalone.
- `historico/` — snapshots acumulados. Não edite à mão.
