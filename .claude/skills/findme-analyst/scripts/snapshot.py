#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
snapshot.py — Registra um snapshot diário de um local no histórico do skill.

Uso:
    python snapshot.py "<historico-dir>" "<snapshot.json>"

O snapshot.json precisa conter ao menos {data, local, ok, parcial, nao_feita,
total, pct_cumprimento}. Os demais campos (dados_inconclusivos, cruzamento,
top_falhas_modelos, etc.) são opcionais mas recomendados — eles que fazem o
histórico ser útil pra detectar padrões.

O arquivo final fica em <historico-dir>/<data>/<slug>.json onde <slug> é o
nome do local em lowercase com underscores, sem acentos e sem "CONDOMÍNIO ".

Se o snapshot daquele local/dia já existir, é sobrescrito (a análise mais
recente do dia ganha).

Imprime na stdout o caminho final escrito, ou {"erro": "..."} em caso de
falha (e sai com código 1).
"""
import sys
import json
import re
import unicodedata
from datetime import date
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def _slug(nome):
    s = unicodedata.normalize("NFD", nome or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"^\s*CONDOMINIO\s+", "", s, flags=re.I).strip()
    s = re.sub(r"[^A-Za-z0-9]+", "_", s.lower()).strip("_")
    return s


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"erro": "uso: python snapshot.py <historico-dir> <snapshot.json>"}))
        sys.exit(1)

    hist_dir = Path(sys.argv[1])
    snap_path = Path(sys.argv[2])

    if not snap_path.exists():
        print(json.dumps({"erro": f"snapshot nao encontrado: {snap_path}"}))
        sys.exit(1)

    try:
        snap = json.loads(snap_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        print(json.dumps({"erro": f"json invalido: {e}"}))
        sys.exit(1)

    required = ["data", "local", "ok", "parcial", "nao_feita", "total"]
    faltando = [k for k in required if k not in snap]
    if faltando:
        print(json.dumps({"erro": f"campos obrigatorios faltando: {faltando}"}))
        sys.exit(1)

    try:
        date.fromisoformat(snap["data"])
    except Exception:
        print(json.dumps({"erro": f"campo 'data' precisa estar em YYYY-MM-DD: {snap['data']!r}"}))
        sys.exit(1)

    slug = _slug(snap["local"])
    if not slug:
        print(json.dumps({"erro": f"nao consegui gerar slug a partir de 'local': {snap['local']!r}"}))
        sys.exit(1)

    # Calcula pct_cumprimento se não veio no snapshot
    if "pct_cumprimento" not in snap and snap["total"]:
        snap["pct_cumprimento"] = round(100.0 * snap["ok"] / snap["total"], 1)

    # Flag de "dados inconclusivos" — se total é alto e ok+parcial = 0,
    # quase certamente é artefato de dados, não falha real.
    if "dados_inconclusivos" not in snap:
        snap["dados_inconclusivos"] = bool(
            snap["total"] >= 10 and snap["ok"] == 0 and snap["parcial"] == 0)

    snap.setdefault("slug", slug)

    out_dir = hist_dir / snap["data"]
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{slug}.json"

    out_path.write_text(json.dumps(snap, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    print(json.dumps({"ok": True, "caminho": str(out_path)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
