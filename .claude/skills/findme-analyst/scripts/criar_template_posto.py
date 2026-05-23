#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
criar_template_posto.py — Cria um arquivo postos/<slug>.json vazio para um
local que ainda não tem registro de avulsas esperadas.

Uso:
    python criar_template_posto.py "<postos-dir>" "<nome do local>"

Se o arquivo já existe, não sobrescreve (avisa e sai). O template criado tem
um único posto exemplo (Limpeza) com uma atividade exemplo comentada via
"_instrucao" — o usuário precisa editar para refletir o que está cadastrado
no portal FindMe.

Imprime na stdout o caminho final escrito, ou {"erro": "..."} em caso de
falha (e sai com código 1).
"""
import sys
import json
import re
import unicodedata
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


TEMPLATE = {
    "_instrucao": (
        "Edite as atividades deste local. Cada entrada tem: modelo (nome exato "
        "como aparece no relatório FindMe), dias (subset de "
        "['Dom','Seg','Ter','Qua','Qui','Sex','Sab']) e vezes (quantas vezes "
        "por dia). Remova este _instrucao quando terminar."
    ),
    "local": "",
    "postos": [
        {
            "posto": "Limpeza",
            "op_tipo": "Limpeza",
            "atividades": []
        }
    ]
}


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"erro": "uso: python criar_template_posto.py <postos-dir> <nome do local>"}))
        sys.exit(1)

    postos_dir = Path(sys.argv[1])
    nome_local = sys.argv[2].strip()

    if not nome_local:
        print(json.dumps({"erro": "nome do local vazio"}))
        sys.exit(1)

    slug = _slug(nome_local)
    if not slug:
        print(json.dumps({"erro": f"nao consegui gerar slug a partir de {nome_local!r}"}))
        sys.exit(1)

    postos_dir.mkdir(parents=True, exist_ok=True)
    out_path = postos_dir / f"{slug}.json"

    if out_path.exists():
        print(json.dumps({
            "erro": f"arquivo ja existe: {out_path}",
            "caminho": str(out_path),
            "ja_existia": True,
        }, ensure_ascii=False))
        sys.exit(1)

    template = dict(TEMPLATE)
    template["local"] = nome_local
    out_path.write_text(json.dumps(template, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    print(json.dumps({"ok": True, "caminho": str(out_path), "slug": slug},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
