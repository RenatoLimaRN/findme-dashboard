#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
notify_failure.py — Dispara email de alerta de falha reaproveitando o
email_config.json existente.

Uso:
    python vps/notify_failure.py \
        --assunto "FALHA análise diária — 2026-05-25" \
        --log logs/run_diario_2026-05-25.log

O corpo do email contém as últimas 80 linhas do log, em monoespaçado.
Se email_config.json não existir, imprime na stderr e sai 1 (sem
quebrar o wrapper bash).
"""
from __future__ import annotations

import argparse
import json
import smtplib
import socket
import sys
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent  # repo root
EMAIL_CFG = ROOT / "email_config.json"

TAIL_LINHAS = 80


def carregar_cfg() -> dict | None:
    if not EMAIL_CFG.exists():
        print(f"[notify_failure] email_config.json nao encontrado em {EMAIL_CFG}",
              file=sys.stderr)
        return None
    try:
        return json.loads(EMAIL_CFG.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[notify_failure] erro ao ler email_config.json: {e}",
              file=sys.stderr)
        return None


def tail_log(log_path: Path, n: int = TAIL_LINHAS) -> str:
    if not log_path.exists():
        return "(log não encontrado)"
    try:
        linhas = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception as e:
        return f"(erro ao ler log: {e})"
    return "\n".join(linhas[-n:])


def montar_email(cfg: dict, assunto: str, log_tail: str, log_path: Path) -> MIMEMultipart:
    from_name = cfg.get("from_name", "FindMe Analyst")
    user = cfg["user"]
    destinatarios = cfg["to"] if isinstance(cfg["to"], list) else [cfg["to"]]
    host = socket.gethostname()
    agora = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")

    msg = MIMEMultipart("alternative")
    msg["From"] = f"{from_name} <{user}>"
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"[FindMe] {assunto}"

    texto = (
        f"Falha na automação diária do FindMe.\n\n"
        f"  Host  : {host}\n"
        f"  Quando: {agora}\n"
        f"  Log   : {log_path}\n\n"
        f"─── últimas {TAIL_LINHAS} linhas do log ────────────────────────\n"
        f"{log_tail}\n"
        f"──────────────────────────────────────────────────────\n\n"
        f"Pra investigar: ssh {host} e ler o arquivo de log inteiro.\n"
        f"Pra rodar manualmente: cd ~/findme-dashboard && bash vps/run_diario.sh\n"
    )

    html = f"""\
<html><body style="font-family:system-ui,sans-serif">
<h2 style="color:#c0392b">⚠️ Falha na automação diária do FindMe</h2>
<table style="border-collapse:collapse">
  <tr><td><b>Host</b></td><td>{host}</td></tr>
  <tr><td><b>Quando</b></td><td>{agora}</td></tr>
  <tr><td><b>Log</b></td><td><code>{log_path}</code></td></tr>
</table>
<h3>Últimas {TAIL_LINHAS} linhas do log</h3>
<pre style="background:#1e1e1e;color:#e0e0e0;padding:12px;border-radius:6px;
  font-size:12px;overflow-x:auto;white-space:pre-wrap">{log_tail}</pre>
<p>Pra investigar: <code>ssh {host}</code> e ler o log inteiro.<br>
Pra rodar manualmente: <code>cd ~/findme-dashboard &amp;&amp; bash vps/run_diario.sh</code></p>
</body></html>"""

    msg.attach(MIMEText(texto, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    return msg


def enviar(cfg: dict, msg: MIMEMultipart) -> bool:
    host = cfg.get("smtp_host", "smtp.gmail.com")
    port = int(cfg.get("smtp_port", 587))
    user = cfg["user"]
    pwd = cfg["password"]
    destinatarios = cfg["to"] if isinstance(cfg["to"], list) else [cfg["to"]]
    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo()
            s.starttls()
            s.login(user, pwd)
            s.sendmail(user, destinatarios, msg.as_string())
        return True
    except Exception as e:
        print(f"[notify_failure] erro SMTP: {e}", file=sys.stderr)
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--assunto", required=True)
    ap.add_argument("--log", type=Path, default=Path("/dev/null"))
    args = ap.parse_args()

    cfg = carregar_cfg()
    if cfg is None:
        sys.exit(1)

    log_tail = tail_log(args.log)
    msg = montar_email(cfg, args.assunto, log_tail, args.log)
    ok = enviar(cfg, msg)
    if ok:
        print(f"[notify_failure] email enviado: {args.assunto}")
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
