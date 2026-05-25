#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gera PROPOSTA_CADASTRO.md a partir do output JSON do popular_postos.py."""
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

d = json.load(sys.stdin)
out = []
out.append("# Proposta de cadastro automático — revisão")
out.append("")
out.append(f"_Gerado a partir de **{d['totais']['locais_no_historico']} locais** no histórico (~135 dias de dados)._")
out.append(f"_Threshold: modelo precisa aparecer em **≥2 dias distintos** pra ser sugerido._")
out.append("")
n_propostas = len([r for r in d["detalhes"] if "postos" in r])
out.append(f"**Resumo:** {n_propostas} locais com proposta. Modo: **{d['modo']}**.")
out.append("")
out.append("---")
out.append("")
out.append("## Como revisar")
out.append("")
out.append("Pra cada local abaixo:")
out.append("- ✅ Tudo OK? Aplica direto.")
out.append("- ✏️ Algum modelo NÃO faz sentido (ex.: avulsa que apareceu poucas vezes)? Abre `postos/<slug>.json` depois de aplicar e remove a entrada.")
out.append("- ✏️ Algum dia da semana errado? Edita o array `dias`.")
out.append("- ✏️ Vezes/dia errado? Edita o campo `vezes`.")
out.append("- ⚠️ Posto inferido errado (ex.: foi pra 'Outros' mas é 'Limpeza')? Move a atividade pra outro objeto `posto`.")
out.append("")
out.append("---")
out.append("")

for r in d["detalhes"]:
    if "postos" not in r:
        out.append(f"## ⚠ {r['nome']}")
        out.append("")
        out.append(f"**{r['acao']}** — {r.get('propostas', 0)} atividade(s) propostas.")
        out.append("")
        continue
    total = sum(len(p["atividades"]) for p in r["postos"])
    out.append(f"## ✓ {r['nome']} ({total} atividade(s))")
    out.append("")
    for p in r["postos"]:
        out.append(f"**Posto: {p['posto']}** (op_tipo: `{p['op_tipo']}`)")
        out.append("")
        out.append("| Modelo | Dias | Vezes/dia | Confiança |")
        out.append("|---|---|---|---|")
        for a in p["atividades"]:
            dias = " ".join(a["dias"])
            out.append(f"| {a['modelo']} | {dias} | {a['vezes']} | {a['_baseado_em']} |")
        out.append("")

Path("PROPOSTA_CADASTRO.md").write_text("\n".join(out), encoding="utf-8")
print(f"Escrito: PROPOSTA_CADASTRO.md ({len(chr(10).join(out))} chars)")
