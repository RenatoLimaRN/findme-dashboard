#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ci_setup.py — Materializa config.json e email_config.json a partir de
variáveis de ambiente (secrets do GitHub Actions). Roda no início do
workflow, antes do analise_diaria.py.

Variáveis esperadas (configuradas em Settings → Secrets do repo no GitHub):
    FINDME_PASSWORD   — senha do FindMe (usada por findme_programacao.py)
    SMTP_USER         — Gmail do remetente (ex.: voce@gmail.com)
    SMTP_PASSWORD     — app password do Gmail (16 chars)
    EMAIL_TO          — destinatários separados por vírgula

Opcional:
    FINDME_EMAIL      — sobrescreve email do config.json se quiser
    SMTP_HOST         — default smtp.gmail.com
    SMTP_PORT         — default 587
    EMAIL_FROM_NAME   — default "FindMe Analyst"
"""
import json
import os
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

ROOT = Path(__file__).resolve().parent


def _need(var):
    v = os.environ.get(var, "").strip()
    if not v:
        print(f"ERRO: variavel de ambiente {var} nao definida", file=sys.stderr)
        sys.exit(1)
    return v


def main():
    # config.json — credenciais do FindMe (locations e cutoff vêm do template)
    tpl = ROOT / "config.json.template"
    if not tpl.exists():
        print(f"ERRO: {tpl} nao encontrado", file=sys.stderr)
        sys.exit(1)
    cfg = json.loads(tpl.read_text(encoding="utf-8"))
    cfg["password"] = _need("FINDME_PASSWORD")
    if os.environ.get("FINDME_EMAIL"):
        cfg["email"] = os.environ["FINDME_EMAIL"]
    (ROOT / "config.json").write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: config.json escrito ({len(cfg.get('locations', []))} locais)")

    # email_config.json — credenciais SMTP
    email_cfg = {
        "smtp_host": os.environ.get("SMTP_HOST", "smtp.gmail.com"),
        "smtp_port": int(os.environ.get("SMTP_PORT", "587")),
        "user": _need("SMTP_USER"),
        "password": _need("SMTP_PASSWORD"),
        "from_name": os.environ.get("EMAIL_FROM_NAME", "FindMe Analyst"),
        "to": [e.strip() for e in _need("EMAIL_TO").split(",") if e.strip()],
    }
    (ROOT / "email_config.json").write_text(
        json.dumps(email_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: email_config.json escrito ({len(email_cfg['to'])} destinatarios)")


if __name__ == "__main__":
    main()
